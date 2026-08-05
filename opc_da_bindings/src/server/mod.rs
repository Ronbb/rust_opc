//! Safe Rust service contracts and COM adapters for OPC Data Access.

use std::collections::HashMap;
use std::ptr;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::{Arc, Mutex};

use opc_classic_types::{ComObject, Error, ErrorCode, Result, Timestamp, Value, ValueType};
use opc_classic_utils::server::{borrow_input, catch_ffi, initialize_output};
use opc_classic_utils::{Cleanup, CoTaskMemArrayBuilder, DropElements, NoCleanup, OwnedPwstr};
use windows::Win32::Foundation::{E_INVALIDARG, E_NOTIMPL, E_OUTOFMEMORY, E_POINTER, E_UNEXPECTED};
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows::Win32::System::Variant::VARIANT;
use windows_core::{
    Error as AbiError, GUID, HRESULT, IUnknown, Interface, OutRef, PCWSTR, PWSTR,
    Result as AbiResult,
};

use crate::abi::{
    error_code_from_abi, error_code_to_abi, object_from_interface, timestamp_to_abi, to_abi_error,
    value_from_abi, value_to_abi,
};
use crate::client::{
    ClientItemHandle, DataSource, GroupOptions, GroupState, ItemSpec, ServerItemHandle,
    ServerStatus,
};
use crate::{
    IOPCGroupStateMgt, IOPCGroupStateMgt_Impl, IOPCItemMgt, IOPCItemMgt_Impl, IOPCServer,
    IOPCServer_Impl, IOPCSyncIO, IOPCSyncIO_Impl, tagOPCDATASOURCE, tagOPCENUMSCOPE, tagOPCITEMDEF,
    tagOPCITEMRESULT, tagOPCITEMSTATE,
};

#[derive(Clone, Debug)]
pub struct ServerSample {
    pub client_handle: ClientItemHandle,
    pub timestamp: Timestamp,
    pub quality: u16,
    pub value: Value,
}

#[derive(Clone, Debug)]
pub struct ServerAddedItem {
    pub server_handle: ServerItemHandle,
    pub canonical_data_type: ValueType,
    pub access_rights: u32,
    pub blob: Vec<u8>,
}

#[derive(Clone, Debug)]
pub struct ServerItemResult<T> {
    pub result: std::result::Result<T, ErrorCode>,
}

impl<T> ServerItemResult<T> {
    pub fn success(value: T) -> Self {
        Self { result: Ok(value) }
    }

    pub fn failure(code: ErrorCode) -> Self {
        Self { result: Err(code) }
    }
}

pub trait DaGroupService: Send + Sync + 'static {
    fn add_items(&self, items: &[ItemSpec]) -> Result<Vec<ServerItemResult<ServerAddedItem>>>;

    fn validate_items(
        &self,
        items: &[ItemSpec],
        _update_blob: bool,
    ) -> Result<Vec<ServerItemResult<ServerAddedItem>>> {
        self.add_items(items)
    }

    fn remove_items(&self, handles: &[ServerItemHandle]) -> Result<Vec<ServerItemResult<()>>>;

    fn set_active(
        &self,
        handles: &[ServerItemHandle],
        active: bool,
    ) -> Result<Vec<ServerItemResult<()>>>;

    fn set_client_handles(
        &self,
        handles: &[(ServerItemHandle, ClientItemHandle)],
    ) -> Result<Vec<ServerItemResult<()>>>;

    fn set_data_types(
        &self,
        handles: &[(ServerItemHandle, ValueType)],
    ) -> Result<Vec<ServerItemResult<()>>>;

    fn read(
        &self,
        source: DataSource,
        handles: &[ServerItemHandle],
    ) -> Result<Vec<ServerItemResult<ServerSample>>>;

    fn write(&self, values: &[(ServerItemHandle, Value)]) -> Result<Vec<ServerItemResult<()>>>;
}

pub trait DaService: Send + Sync + 'static {
    fn status(&self) -> Result<ServerStatus>;
    fn add_group(&self, options: GroupOptions) -> Result<Arc<dyn DaGroupService>>;
    fn remove_group(&self, server_handle: u32, force: bool) -> Result<()>;
    fn error_string(&self, error: ErrorCode, locale: u32) -> Result<String>;
}

#[derive(Clone)]
struct GroupEntry {
    name: String,
    object: IUnknown,
}

#[derive(Clone)]
pub struct DaServer {
    inner: IOPCServer,
    service: Arc<dyn DaService>,
}

impl DaServer {
    pub fn new(service: Arc<dyn DaService>) -> Self {
        let inner: IOPCServer = DaServerAdapter::new(service.clone()).into();
        Self { inner, service }
    }

    pub fn service(&self) -> Arc<dyn DaService> {
        self.service.clone()
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }
}

#[windows_core::implement(IOPCServer)]
struct DaServerAdapter {
    service: Arc<dyn DaService>,
    groups: Mutex<HashMap<u32, GroupEntry>>,
    next_group_handle: AtomicU32,
}

impl DaServerAdapter {
    fn new(service: Arc<dyn DaService>) -> Self {
        Self {
            service,
            groups: Mutex::new(HashMap::new()),
            next_group_handle: AtomicU32::new(1),
        }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCServer_Impl for DaServerAdapter_Impl {
    fn AddGroup(
        &self,
        name: &PCWSTR,
        active: windows_core::BOOL,
        requested_update_rate: u32,
        client_handle: u32,
        time_bias: *const i32,
        percent_deadband: *const f32,
        locale: u32,
        server_handle: *mut u32,
        revised_update_rate: *mut u32,
        iid: *const GUID,
        output: OutRef<IUnknown>,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe {
                initialize_output(server_handle)?;
                initialize_output(revised_update_rate)?;
            }
            if output.is_null() {
                return Err(code_error(E_POINTER));
            }
            // SAFETY: `OutRef<IUnknown>` is a transparent wrapper over the ABI
            // output pointer. Initialize it without consuming `output`, which
            // is needed again for the successful interface transfer.
            let output_ptr: *mut *mut core::ffi::c_void =
                unsafe { std::mem::transmute_copy(&output) };
            unsafe { initialize_output(output_ptr)? };
            if iid.is_null() {
                return Err(code_error(E_POINTER));
            }
            if name.is_null() {
                return Err(Error::null_pointer("null group name"));
            }
            let name = unsafe { name.to_string() }
                .map_err(|_| Error::invalid_argument("invalid UTF-16 group name"))?;
            let options = GroupOptions {
                name: name.clone(),
                active: active.as_bool(),
                requested_update_rate,
                client_handle,
                time_bias: unsafe { time_bias.as_ref().copied() },
                percent_deadband: unsafe { percent_deadband.as_ref().copied() },
                locale,
            };
            let service = self.service.add_group(options.clone())?;
            let handle = next_nonzero(&self.next_group_handle)?;
            let state = GroupState {
                update_rate: options.requested_update_rate,
                active: options.active,
                name: name.clone(),
                time_bias: options.time_bias.unwrap_or_default(),
                percent_deadband: options.percent_deadband.unwrap_or_default(),
                locale: options.locale,
                client_handle: options.client_handle,
                server_handle: handle,
            };
            let group: IUnknown = DaGroupAdapter::new(service, state).into();
            let requested = query_interface(&group, unsafe { &*iid })?;
            self.groups
                .lock()
                .map_err(|_| code_error(E_UNEXPECTED))?
                .insert(
                    handle,
                    GroupEntry {
                        name,
                        object: group,
                    },
                );
            output
                .write(Some(requested))
                .map_err(crate::abi::from_abi_error)?;
            unsafe {
                server_handle.write(handle);
                revised_update_rate.write(requested_update_rate);
            }
            Ok(())
        })
    }

    fn GetErrorString(&self, error: HRESULT, locale: u32) -> AbiResult<PWSTR> {
        serve(|| {
            let value = self
                .service
                .error_string(error_code_from_abi(error), locale)?;
            Ok(PWSTR(OwnedPwstr::new(value)?.into_raw()))
        })
    }

    fn GetGroupByName(&self, name: &PCWSTR, iid: *const GUID) -> AbiResult<IUnknown> {
        serve(|| {
            if iid.is_null() {
                return Err(code_error(E_POINTER));
            }
            if name.is_null() {
                return Err(Error::null_pointer("null group name"));
            }
            let name = unsafe { name.to_string() }
                .map_err(|_| Error::invalid_argument("invalid UTF-16 group name"))?;
            let groups = self.groups.lock().map_err(|_| code_error(E_UNEXPECTED))?;
            let group = groups
                .values()
                .find(|group| group.name == name)
                .ok_or_else(|| code_error(E_INVALIDARG))?;
            query_interface(&group.object, unsafe { &*iid })
        })
    }

    fn GetStatus(&self) -> AbiResult<*mut crate::tagOPCSERVERSTATUS> {
        serve(|| {
            let status = self.service.status()?;
            let vendor = OwnedPwstr::new(status.vendor_info)?;
            let mut output = CoTaskMemArrayBuilder::new(1, ServerStatusCleanup)?;
            let value = crate::tagOPCSERVERSTATUS {
                ftStartTime: timestamp_to_abi(status.start_time),
                ftCurrentTime: timestamp_to_abi(status.current_time),
                ftLastUpdateTime: timestamp_to_abi(status.last_update_time),
                dwServerState: crate::tagOPCSERVERSTATE(status.state.raw()),
                dwGroupCount: status.group_count,
                dwBandWidth: status.bandwidth,
                wMajorVersion: status.version.0,
                wMinorVersion: status.version.1,
                wBuildNumber: status.version.2,
                wReserved: 0,
                szVendorInfo: PWSTR(vendor.into_raw()),
            };
            if let Err(mut value) = output.push(value) {
                unsafe { ServerStatusCleanup.cleanup(&mut value, 1) };
                return Err(code_error(E_UNEXPECTED));
            }
            let (ptr, _) = output.finish()?.into_raw_parts();
            Ok(ptr)
        })
    }

    fn RemoveGroup(&self, server_handle: u32, force: windows_core::BOOL) -> AbiResult<()> {
        serve(|| {
            self.service.remove_group(server_handle, force.as_bool())?;
            self.groups
                .lock()
                .map_err(|_| code_error(E_UNEXPECTED))?
                .remove(&server_handle);
            Ok(())
        })
    }

    fn CreateGroupEnumerator(
        &self,
        _scope: tagOPCENUMSCOPE,
        _iid: *const GUID,
    ) -> AbiResult<IUnknown> {
        Err(AbiError::from_hresult(E_NOTIMPL))
    }
}

struct ServerStatusCleanup;

// SAFETY: The status contains one task-allocated vendor string.
unsafe impl Cleanup<crate::tagOPCSERVERSTATUS> for ServerStatusCleanup {
    unsafe fn cleanup(&mut self, ptr: *mut crate::tagOPCSERVERSTATUS, initialized: usize) {
        for index in 0..initialized {
            let status = unsafe { &mut *ptr.add(index) };
            if !status.szVendorInfo.is_null() {
                unsafe { CoTaskMemFree(Some(status.szVendorInfo.0.cast())) };
                status.szVendorInfo = PWSTR::null();
            }
        }
    }
}

#[windows_core::implement(IOPCItemMgt, IOPCSyncIO, IOPCGroupStateMgt)]
struct DaGroupAdapter {
    service: Arc<dyn DaGroupService>,
    state: Mutex<GroupState>,
}

impl DaGroupAdapter {
    fn new(service: Arc<dyn DaGroupService>, state: GroupState) -> Self {
        Self {
            service,
            state: Mutex::new(state),
        }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCItemMgt_Impl for DaGroupAdapter_Impl {
    fn AddItems(
        &self,
        count: u32,
        items: *const tagOPCITEMDEF,
        results: *mut *mut tagOPCITEMRESULT,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe {
                initialize_output(results)?;
                initialize_output(errors)?;
            }
            let items = decode_item_specs(unsafe { borrow_input(items, count)? })?;
            let values = self.service.add_items(&items)?;
            write_item_results(values, count as usize, results, errors)
        })
    }

    fn ValidateItems(
        &self,
        count: u32,
        items: *const tagOPCITEMDEF,
        update_blob: windows_core::BOOL,
        results: *mut *mut tagOPCITEMRESULT,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe {
                initialize_output(results)?;
                initialize_output(errors)?;
            }
            let items = decode_item_specs(unsafe { borrow_input(items, count)? })?;
            let values = self.service.validate_items(&items, update_blob.as_bool())?;
            write_item_results(values, count as usize, results, errors)
        })
    }

    fn RemoveItems(
        &self,
        count: u32,
        handles: *const u32,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(errors)? };
            let handles = decode_handles(unsafe { borrow_input(handles, count)? });
            let values = self.service.remove_items(&handles)?;
            write_unit_results(values, count as usize, errors)
        })
    }

    fn SetActiveState(
        &self,
        count: u32,
        handles: *const u32,
        active: windows_core::BOOL,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(errors)? };
            let handles = decode_handles(unsafe { borrow_input(handles, count)? });
            let values = self.service.set_active(&handles, active.as_bool())?;
            write_unit_results(values, count as usize, errors)
        })
    }

    fn SetClientHandles(
        &self,
        count: u32,
        server_handles: *const u32,
        client_handles: *const u32,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(errors)? };
            let server_handles = unsafe { borrow_input(server_handles, count)? };
            let client_handles = unsafe { borrow_input(client_handles, count)? };
            let pairs: Vec<_> = server_handles
                .iter()
                .zip(client_handles)
                .map(|(server, client)| (ServerItemHandle(*server), ClientItemHandle(*client)))
                .collect();
            let values = self.service.set_client_handles(&pairs)?;
            write_unit_results(values, count as usize, errors)
        })
    }

    fn SetDatatypes(
        &self,
        count: u32,
        handles: *const u32,
        data_types: *const u16,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(errors)? };
            let handles = unsafe { borrow_input(handles, count)? };
            let data_types = unsafe { borrow_input(data_types, count)? };
            let pairs: Vec<_> = handles
                .iter()
                .zip(data_types)
                .map(|(handle, data_type)| {
                    (ServerItemHandle(*handle), ValueType::from_raw(*data_type))
                })
                .collect();
            let values = self.service.set_data_types(&pairs)?;
            write_unit_results(values, count as usize, errors)
        })
    }

    fn CreateEnumerator(&self, _iid: *const GUID) -> AbiResult<IUnknown> {
        Err(AbiError::from_hresult(E_NOTIMPL))
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCSyncIO_Impl for DaGroupAdapter_Impl {
    fn Read(
        &self,
        source: tagOPCDATASOURCE,
        count: u32,
        handles: *const u32,
        values: *mut *mut tagOPCITEMSTATE,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe {
                initialize_output(values)?;
                initialize_output(errors)?;
            }
            let source = if source == crate::OPC_DS_CACHE {
                DataSource::Cache
            } else if source == crate::OPC_DS_DEVICE {
                DataSource::Device
            } else {
                return Err(code_error(E_INVALIDARG));
            };
            let handles = decode_handles(unsafe { borrow_input(handles, count)? });
            let returned = self.service.read(source, &handles)?;
            ensure_batch_len(returned.len(), count as usize)?;
            let mut states = CoTaskMemArrayBuilder::new(count as usize, DropElements)?;
            let mut item_errors = CoTaskMemArrayBuilder::new(count as usize, NoCleanup)?;
            let mut has_error = false;
            for result in returned {
                match result.result {
                    Ok(sample) => {
                        states
                            .push(tagOPCITEMSTATE {
                                hClient: sample.client_handle.0,
                                ftTimeStamp: timestamp_to_abi(sample.timestamp),
                                wQuality: sample.quality,
                                wReserved: 0,
                                vDataValue: value_to_abi(&sample.value)?,
                            })
                            .map_err(|_| code_error(E_UNEXPECTED))?;
                        item_errors
                            .push(HRESULT(0))
                            .map_err(|_| code_error(E_UNEXPECTED))?;
                    }
                    Err(error) => {
                        has_error = true;
                        states
                            .push(tagOPCITEMSTATE::default())
                            .map_err(|_| code_error(E_UNEXPECTED))?;
                        item_errors
                            .push(error_code_to_abi(error))
                            .map_err(|_| code_error(E_UNEXPECTED))?;
                    }
                }
            }
            let states = states.finish()?;
            let item_errors = item_errors.finish()?;
            let (state_ptr, _) = states.into_raw_parts();
            let (error_ptr, _) = item_errors.into_raw_parts();
            unsafe {
                values.write(state_ptr);
                errors.write(error_ptr);
            }
            batch_status(has_error)
        })
    }

    fn Write(
        &self,
        count: u32,
        handles: *const u32,
        values: *const VARIANT,
        errors: *mut *mut HRESULT,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(errors)? };
            let handles = unsafe { borrow_input(handles, count)? };
            let values = unsafe { borrow_input(values, count)? };
            let owned: Vec<_> = handles
                .iter()
                .zip(values)
                .map(|(handle, value)| Ok((ServerItemHandle(*handle), value_from_abi(value)?)))
                .collect::<Result<_>>()?;
            let returned = self.service.write(&owned)?;
            write_unit_results(returned, count as usize, errors)
        })
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCGroupStateMgt_Impl for DaGroupAdapter_Impl {
    fn GetState(
        &self,
        update_rate: *mut u32,
        active: *mut windows_core::BOOL,
        name: *mut PWSTR,
        time_bias: *mut i32,
        percent_deadband: *mut f32,
        locale: *mut u32,
        client_handle: *mut u32,
        server_handle: *mut u32,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe {
                initialize_output(update_rate)?;
                initialize_output(active)?;
                initialize_output(name)?;
                initialize_output(time_bias)?;
                initialize_output(percent_deadband)?;
                initialize_output(locale)?;
                initialize_output(client_handle)?;
                initialize_output(server_handle)?;
            }
            let state = self
                .state
                .lock()
                .map_err(|_| code_error(E_UNEXPECTED))?
                .clone();
            let name_value = OwnedPwstr::new(state.name)?;
            unsafe {
                update_rate.write(state.update_rate);
                active.write(state.active.into());
                name.write(PWSTR(name_value.into_raw()));
                time_bias.write(state.time_bias);
                percent_deadband.write(state.percent_deadband);
                locale.write(state.locale);
                client_handle.write(state.client_handle);
                server_handle.write(state.server_handle);
            }
            Ok(())
        })
    }

    fn SetState(
        &self,
        requested_update_rate: *const u32,
        revised_update_rate: *mut u32,
        active: *const windows_core::BOOL,
        time_bias: *const i32,
        percent_deadband: *const f32,
        locale: *const u32,
        client_handle: *const u32,
    ) -> AbiResult<()> {
        serve(|| {
            unsafe { initialize_output(revised_update_rate)? };
            let mut state = self.state.lock().map_err(|_| code_error(E_UNEXPECTED))?;
            if let Some(value) = unsafe { requested_update_rate.as_ref() } {
                state.update_rate = *value;
            }
            if let Some(value) = unsafe { active.as_ref() } {
                state.active = value.as_bool();
            }
            if let Some(value) = unsafe { time_bias.as_ref() } {
                state.time_bias = *value;
            }
            if let Some(value) = unsafe { percent_deadband.as_ref() } {
                state.percent_deadband = *value;
            }
            if let Some(value) = unsafe { locale.as_ref() } {
                state.locale = *value;
            }
            if let Some(value) = unsafe { client_handle.as_ref() } {
                state.client_handle = *value;
            }
            unsafe { revised_update_rate.write(state.update_rate) };
            Ok(())
        })
    }

    fn SetName(&self, name: &PCWSTR) -> AbiResult<()> {
        serve(|| {
            if name.is_null() {
                return Err(Error::null_pointer("null group name"));
            }
            self.state
                .lock()
                .map_err(|_| code_error(E_UNEXPECTED))?
                .name = unsafe { name.to_string() }
                .map_err(|_| Error::invalid_argument("invalid UTF-16 group name"))?;
            Ok(())
        })
    }

    fn CloneGroup(&self, _name: &PCWSTR, _iid: *const GUID) -> AbiResult<IUnknown> {
        Err(AbiError::from_hresult(E_NOTIMPL))
    }
}

struct ItemResultCleanup;

// SAFETY: OPCITEMRESULT owns one optional task-allocated blob.
unsafe impl Cleanup<tagOPCITEMRESULT> for ItemResultCleanup {
    unsafe fn cleanup(&mut self, ptr: *mut tagOPCITEMRESULT, initialized: usize) {
        for index in 0..initialized {
            let value = unsafe { &mut *ptr.add(index) };
            if !value.pBlob.is_null() {
                unsafe { CoTaskMemFree(Some(value.pBlob.cast())) };
                value.pBlob = ptr::null_mut();
                value.dwBlobSize = 0;
            }
        }
    }
}

fn decode_item_specs(values: &[tagOPCITEMDEF]) -> Result<Vec<ItemSpec>> {
    values
        .iter()
        .map(|value| {
            if value.dwBlobSize != 0 && value.pBlob.is_null() {
                return Err(code_error(E_INVALIDARG));
            }
            let blob = if value.dwBlobSize == 0 {
                Vec::new()
            } else {
                unsafe { std::slice::from_raw_parts(value.pBlob, value.dwBlobSize as usize) }
                    .to_vec()
            };
            Ok(ItemSpec {
                item_id: pwstr_string(value.szItemID)?,
                access_path: pwstr_string(value.szAccessPath)?,
                active: value.bActive.as_bool(),
                client_handle: ClientItemHandle(value.hClient),
                requested_data_type: ValueType::from_raw(value.vtRequestedDataType),
                blob,
            })
        })
        .collect()
}

fn pwstr_string(value: PWSTR) -> Result<String> {
    if value.is_null() {
        Ok(String::new())
    } else {
        unsafe { value.to_string() }.map_err(|_| Error::invalid_argument("invalid UTF-16 string"))
    }
}

fn decode_handles(values: &[u32]) -> Vec<ServerItemHandle> {
    values.iter().copied().map(ServerItemHandle).collect()
}

fn write_item_results(
    values: Vec<ServerItemResult<ServerAddedItem>>,
    expected: usize,
    results: *mut *mut tagOPCITEMRESULT,
    errors: *mut *mut HRESULT,
) -> Result<()> {
    unsafe {
        initialize_output(results)?;
        initialize_output(errors)?;
    }
    ensure_batch_len(values.len(), expected)?;
    let mut result_values = CoTaskMemArrayBuilder::new(expected, ItemResultCleanup)?;
    let mut item_errors = CoTaskMemArrayBuilder::new(expected, NoCleanup)?;
    let mut has_error = false;
    for value in values {
        match value.result {
            Ok(value) => {
                let (blob, blob_size) = allocate_blob(&value.blob)?;
                push_item_result(
                    &mut result_values,
                    tagOPCITEMRESULT {
                        hServer: value.server_handle.0,
                        vtCanonicalDataType: value.canonical_data_type.raw(),
                        wReserved: 0,
                        dwAccessRights: value.access_rights,
                        dwBlobSize: blob_size,
                        pBlob: blob,
                    },
                )?;
                item_errors
                    .push(HRESULT(0))
                    .map_err(|_| code_error(E_UNEXPECTED))?;
            }
            Err(error) => {
                has_error = true;
                push_item_result(&mut result_values, tagOPCITEMRESULT::default())?;
                item_errors
                    .push(error_code_to_abi(error))
                    .map_err(|_| code_error(E_UNEXPECTED))?;
            }
        }
    }
    let result_values = result_values.finish()?;
    let item_errors = item_errors.finish()?;
    let (result_ptr, _) = result_values.into_raw_parts();
    let (error_ptr, _) = item_errors.into_raw_parts();
    unsafe {
        results.write(result_ptr);
        errors.write(error_ptr);
    }
    batch_status(has_error)
}

fn push_item_result(
    output: &mut CoTaskMemArrayBuilder<tagOPCITEMRESULT, ItemResultCleanup>,
    value: tagOPCITEMRESULT,
) -> Result<()> {
    if let Err(mut value) = output.push(value) {
        unsafe { ItemResultCleanup.cleanup(&mut value, 1) };
        Err(code_error(E_UNEXPECTED))
    } else {
        Ok(())
    }
}

fn write_unit_results(
    values: Vec<ServerItemResult<()>>,
    expected: usize,
    errors: *mut *mut HRESULT,
) -> Result<()> {
    unsafe { initialize_output(errors)? };
    ensure_batch_len(values.len(), expected)?;
    let mut item_errors = CoTaskMemArrayBuilder::new(expected, NoCleanup)?;
    let mut has_error = false;
    for value in values {
        let error = match value.result {
            Ok(()) => HRESULT(0),
            Err(error) => {
                has_error = true;
                error_code_to_abi(error)
            }
        };
        item_errors
            .push(error)
            .map_err(|_| code_error(E_UNEXPECTED))?;
    }
    let (error_ptr, _) = item_errors.finish()?.into_raw_parts();
    unsafe { errors.write(error_ptr) };
    batch_status(has_error)
}

fn allocate_blob(value: &[u8]) -> Result<(*mut u8, u32)> {
    if value.is_empty() {
        return Ok((ptr::null_mut(), 0));
    }
    let len = u32::try_from(value.len()).map_err(|_| code_error(E_OUTOFMEMORY))?;
    let allocation = unsafe { CoTaskMemAlloc(value.len()) }.cast::<u8>();
    if allocation.is_null() {
        return Err(code_error(E_OUTOFMEMORY));
    }
    unsafe { ptr::copy_nonoverlapping(value.as_ptr(), allocation, value.len()) };
    Ok((allocation, len))
}

fn ensure_batch_len(actual: usize, expected: usize) -> Result<()> {
    if actual == expected {
        Ok(())
    } else {
        Err(code_error(E_UNEXPECTED))
    }
}

fn batch_status(has_error: bool) -> Result<()> {
    if has_error {
        // The generated trampoline converts this `Error` back to its positive
        // success HRESULT, preserving OPC's partial-success status.
        Err(Error::from_code(ErrorCode::PARTIAL_SUCCESS))
    } else {
        Ok(())
    }
}

fn next_nonzero(counter: &AtomicU32) -> Result<u32> {
    let value = counter.fetch_add(1, Ordering::Relaxed);
    if value == 0 || value == u32::MAX {
        Err(code_error(E_OUTOFMEMORY))
    } else {
        Ok(value)
    }
}

fn query_interface(object: &IUnknown, iid: &GUID) -> Result<IUnknown> {
    let mut raw = ptr::null_mut();
    let call = unsafe { object.query(iid, &mut raw) };
    let queried = if raw.is_null() {
        Err(code_error(E_POINTER))
    } else {
        Ok(unsafe { IUnknown::from_raw(raw) })
    };
    call.ok().map_err(crate::abi::from_abi_error)?;
    queried
}

fn serve<T>(f: impl FnOnce() -> Result<T>) -> AbiResult<T> {
    catch_ffi(f).map_err(to_abi_error)
}

fn code_error(code: HRESULT) -> Error {
    Error::from_code(error_code_from_abi(code))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::client::{DaClient, GroupOptions, ServerState};

    struct TestService;
    struct TestGroup;

    impl DaService for TestService {
        fn status(&self) -> Result<ServerStatus> {
            Ok(ServerStatus {
                start_time: Timestamp::default(),
                current_time: Timestamp::default(),
                last_update_time: Timestamp::default(),
                state: ServerState::Running,
                group_count: 0,
                bandwidth: 1,
                version: (1, 2, 3),
                vendor_info: "rust-opc test".into(),
            })
        }

        fn add_group(&self, _options: GroupOptions) -> Result<Arc<dyn DaGroupService>> {
            Ok(Arc::new(TestGroup))
        }

        fn remove_group(&self, _server_handle: u32, _force: bool) -> Result<()> {
            Ok(())
        }

        fn error_string(&self, _error: ErrorCode, _locale: u32) -> Result<String> {
            Ok("test error".into())
        }
    }

    impl DaGroupService for TestGroup {
        fn add_items(&self, items: &[ItemSpec]) -> Result<Vec<ServerItemResult<ServerAddedItem>>> {
            Ok(items
                .iter()
                .enumerate()
                .map(|(index, _)| {
                    ServerItemResult::success(ServerAddedItem {
                        server_handle: ServerItemHandle(index as u32 + 10),
                        canonical_data_type: ValueType::I32,
                        access_rights: 3,
                        blob: vec![1, 2],
                    })
                })
                .collect())
        }

        fn remove_items(&self, handles: &[ServerItemHandle]) -> Result<Vec<ServerItemResult<()>>> {
            Ok(handles
                .iter()
                .map(|_| ServerItemResult::success(()))
                .collect())
        }

        fn set_active(
            &self,
            handles: &[ServerItemHandle],
            _active: bool,
        ) -> Result<Vec<ServerItemResult<()>>> {
            Ok(handles
                .iter()
                .map(|_| ServerItemResult::success(()))
                .collect())
        }

        fn set_client_handles(
            &self,
            handles: &[(ServerItemHandle, ClientItemHandle)],
        ) -> Result<Vec<ServerItemResult<()>>> {
            Ok(handles
                .iter()
                .map(|_| ServerItemResult::success(()))
                .collect())
        }

        fn set_data_types(
            &self,
            handles: &[(ServerItemHandle, ValueType)],
        ) -> Result<Vec<ServerItemResult<()>>> {
            Ok(handles
                .iter()
                .map(|_| ServerItemResult::success(()))
                .collect())
        }

        fn read(
            &self,
            _source: DataSource,
            handles: &[ServerItemHandle],
        ) -> Result<Vec<ServerItemResult<ServerSample>>> {
            Ok(handles
                .iter()
                .map(|_| {
                    ServerItemResult::success(ServerSample {
                        client_handle: ClientItemHandle(7),
                        timestamp: Timestamp::default(),
                        quality: 192,
                        value: Value::I32(42),
                    })
                })
                .collect())
        }

        fn write(&self, values: &[(ServerItemHandle, Value)]) -> Result<Vec<ServerItemResult<()>>> {
            Ok(values
                .iter()
                .map(|_| ServerItemResult::success(()))
                .collect())
        }
    }

    #[test]
    fn client_and_server_adapters_round_trip_owned_values() {
        let apartment = opc_classic_utils::ComApartment::mta().unwrap();
        let server = DaServer::new(Arc::new(TestService));
        let client = DaClient::from_object(&apartment, &server.object()).unwrap();

        let status = client.status().unwrap();
        assert_eq!(status.vendor_info, "rust-opc test");

        let group = client.add_group(GroupOptions::new("test")).unwrap();
        let items = group.add_items(&[ItemSpec::new("item", 7)]).unwrap();
        let item = items[0].as_ref().unwrap();
        assert_eq!(item.server_handle.raw(), 10);
        assert_eq!(item.blob, [1, 2]);

        let values = group
            .read(DataSource::Cache, &[item.server_handle])
            .unwrap();
        assert_eq!(values[0].as_ref().unwrap().quality, 192);
        assert_eq!(values[0].as_ref().unwrap().value, Value::I32(42));
        group.close(false).unwrap();
    }
}
