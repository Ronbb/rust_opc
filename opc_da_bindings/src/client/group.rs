use std::marker::PhantomData;
use std::ptr;

use opc_classic_utils::{Cleanup, CoTaskMemOut, DropElements, NoCleanup, OwnedPwstr, WideCString};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Variant::VARIANT;
use windows_core::{Error, HRESULT, IUnknown, Interface, Result};

use crate::{
    IOPCGroupStateMgt, IOPCItemMgt, IOPCSyncIO, tagOPCITEMDEF, tagOPCITEMRESULT, tagOPCITEMSTATE,
};

use super::{
    AddedItem, ClientItemHandle, DataSource, GroupState, ItemError, ItemSpec, Sample,
    ServerItemHandle, WriteValue,
};

pub(crate) struct ItemResultCleanup;

// SAFETY: OPCITEMRESULT owns only its optional task-allocated blob.
unsafe impl Cleanup<tagOPCITEMRESULT> for ItemResultCleanup {
    unsafe fn cleanup(&mut self, ptr: *mut tagOPCITEMRESULT, initialized: usize) {
        for index in 0..initialized {
            let item = unsafe { &mut *ptr.add(index) };
            if !item.pBlob.is_null() {
                unsafe { CoTaskMemFree(Some(item.pBlob.cast())) };
                item.pBlob = ptr::null_mut();
                item.dwBlobSize = 0;
            }
        }
    }
}

pub struct DaGroup<'apartment> {
    server: crate::IOPCServer,
    object: IUnknown,
    server_handle: Option<u32>,
    remove_on_drop: bool,
    revised_update_rate: u32,
    _apartment: PhantomData<&'apartment opc_classic_utils::ComApartment>,
}

impl<'apartment> DaGroup<'apartment> {
    pub(crate) fn new(
        server: crate::IOPCServer,
        object: IUnknown,
        server_handle: u32,
        remove_on_drop: bool,
        revised_update_rate: u32,
    ) -> Self {
        Self {
            server,
            object,
            server_handle: Some(server_handle),
            remove_on_drop,
            revised_update_rate,
            _apartment: PhantomData,
        }
    }

    pub fn revised_update_rate(&self) -> u32 {
        self.revised_update_rate
    }

    pub fn state(&self) -> Result<GroupState> {
        let state: IOPCGroupStateMgt = self.object.cast()?;
        let mut update_rate = 0;
        let mut active = windows_core::BOOL(0);
        let mut name = windows_core::PWSTR::null();
        let mut time_bias = 0;
        let mut percent_deadband = 0.0;
        let mut locale = 0;
        let mut client_handle = 0;
        let mut server_handle = 0;
        let call = unsafe {
            state.GetState(
                &mut update_rate,
                &mut active,
                &mut name,
                &mut time_bias,
                &mut percent_deadband,
                &mut locale,
                &mut client_handle,
                &mut server_handle,
            )
        };
        let name = unsafe { OwnedPwstr::from_raw(name.0) };
        call?;
        Ok(GroupState {
            update_rate,
            active: active.as_bool(),
            name: name.to_string_lossy(),
            time_bias,
            percent_deadband,
            locale,
            client_handle,
            server_handle,
        })
    }

    pub fn set_name(&self, name: &str) -> Result<()> {
        let name = wide(name)?;
        let state: IOPCGroupStateMgt = self.object.cast()?;
        unsafe { state.SetName(name.as_pcwstr()) }
    }

    pub fn add_items(
        &self,
        specs: &[ItemSpec],
    ) -> Result<Vec<std::result::Result<AddedItem, ItemError>>> {
        validate_blob_lengths(specs)?;
        let item_mgt: IOPCItemMgt = self.object.cast()?;
        let ids = specs
            .iter()
            .map(|spec| wide(&spec.item_id))
            .collect::<Result<Vec<_>>>()?;
        let paths = specs
            .iter()
            .map(|spec| wide(&spec.access_path))
            .collect::<Result<Vec<_>>>()?;
        let definitions: Vec<_> = specs
            .iter()
            .zip(ids.iter().zip(paths.iter()))
            .map(|(spec, (id, path))| tagOPCITEMDEF {
                szAccessPath: windows_core::PWSTR(path.as_pcwstr().0.cast_mut()),
                szItemID: windows_core::PWSTR(id.as_pcwstr().0.cast_mut()),
                bActive: spec.active.into(),
                hClient: spec.client_handle.0,
                dwBlobSize: spec.blob.len() as u32,
                pBlob: blob_ptr(&spec.blob),
                vtRequestedDataType: spec.requested_data_type,
                wReserved: 0,
            })
            .collect();

        let count = count(definitions.len())?;
        let mut results = CoTaskMemOut::<tagOPCITEMRESULT>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            item_mgt.AddItems(
                count,
                definitions.as_ptr(),
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let results = unsafe { results.into_array(specs.len(), ItemResultCleanup) }?;
        let errors = unsafe { errors.into_array(specs.len(), NoCleanup) }?;
        call?;

        Ok(results
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(result, error)| {
                if error.is_err() {
                    Err(ItemError { code: *error })
                } else {
                    let blob = if result.pBlob.is_null() || result.dwBlobSize == 0 {
                        Vec::new()
                    } else {
                        unsafe {
                            std::slice::from_raw_parts(result.pBlob, result.dwBlobSize as usize)
                        }
                        .to_vec()
                    };
                    Ok(AddedItem {
                        server_handle: ServerItemHandle(result.hServer),
                        canonical_data_type: result.vtCanonicalDataType,
                        access_rights: result.dwAccessRights,
                        blob,
                    })
                }
            })
            .collect())
    }

    pub fn validate_items(
        &self,
        specs: &[ItemSpec],
        update_blob: bool,
    ) -> Result<Vec<std::result::Result<AddedItem, ItemError>>> {
        validate_blob_lengths(specs)?;
        let item_mgt: IOPCItemMgt = self.object.cast()?;
        let ids = specs
            .iter()
            .map(|spec| wide(&spec.item_id))
            .collect::<Result<Vec<_>>>()?;
        let paths = specs
            .iter()
            .map(|spec| wide(&spec.access_path))
            .collect::<Result<Vec<_>>>()?;
        let definitions: Vec<_> = specs
            .iter()
            .zip(ids.iter().zip(paths.iter()))
            .map(|(spec, (id, path))| tagOPCITEMDEF {
                szAccessPath: windows_core::PWSTR(path.as_pcwstr().0.cast_mut()),
                szItemID: windows_core::PWSTR(id.as_pcwstr().0.cast_mut()),
                bActive: spec.active.into(),
                hClient: spec.client_handle.0,
                dwBlobSize: spec.blob.len() as u32,
                pBlob: blob_ptr(&spec.blob),
                vtRequestedDataType: spec.requested_data_type,
                wReserved: 0,
            })
            .collect();
        let mut results = CoTaskMemOut::<tagOPCITEMRESULT>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            item_mgt.ValidateItems(
                count(specs.len())?,
                definitions.as_ptr(),
                update_blob,
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let results = unsafe { results.into_array(specs.len(), ItemResultCleanup) }?;
        let errors = unsafe { errors.into_array(specs.len(), NoCleanup) }?;
        call?;
        Ok(results
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(result, error)| {
                if error.is_err() {
                    Err(ItemError { code: *error })
                } else {
                    let blob = if result.pBlob.is_null() || result.dwBlobSize == 0 {
                        Vec::new()
                    } else {
                        unsafe {
                            std::slice::from_raw_parts(result.pBlob, result.dwBlobSize as usize)
                        }
                        .to_vec()
                    };
                    Ok(AddedItem {
                        server_handle: ServerItemHandle(result.hServer),
                        canonical_data_type: result.vtCanonicalDataType,
                        access_rights: result.dwAccessRights,
                        blob,
                    })
                }
            })
            .collect())
    }

    pub fn remove_items(
        &self,
        handles: &[ServerItemHandle],
    ) -> Result<Vec<std::result::Result<(), ItemError>>> {
        let item_mgt: IOPCItemMgt = self.object.cast()?;
        let raw = raw_handles(handles);
        item_errors(handles.len(), |errors| unsafe {
            item_mgt.RemoveItems(count(handles.len())?, raw.as_ptr(), errors)
        })
    }

    pub fn set_items_active(
        &self,
        handles: &[ServerItemHandle],
        active: bool,
    ) -> Result<Vec<std::result::Result<(), ItemError>>> {
        let item_mgt: IOPCItemMgt = self.object.cast()?;
        let raw = raw_handles(handles);
        item_errors(handles.len(), |errors| unsafe {
            item_mgt.SetActiveState(count(handles.len())?, raw.as_ptr(), active, errors)
        })
    }

    pub fn read(
        &self,
        source: DataSource,
        handles: &[ServerItemHandle],
    ) -> Result<Vec<std::result::Result<Sample, ItemError>>> {
        let io: IOPCSyncIO = self.object.cast()?;
        let raw = raw_handles(handles);
        let mut values = CoTaskMemOut::<tagOPCITEMSTATE>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            io.Read(
                source.as_abi(),
                count(handles.len())?,
                raw.as_ptr(),
                values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let values = unsafe { values.into_array(handles.len(), DropElements) }?;
        let errors = unsafe { errors.into_array(handles.len(), NoCleanup) }?;
        call?;
        Ok(values
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(value, error)| {
                if error.is_err() {
                    Err(ItemError { code: *error })
                } else {
                    Ok(Sample {
                        client_handle: ClientItemHandle(value.hClient),
                        timestamp: value.ftTimeStamp,
                        quality: value.wQuality,
                        value: value.vDataValue.clone(),
                    })
                }
            })
            .collect())
    }

    pub fn write(&self, values: &[WriteValue]) -> Result<Vec<std::result::Result<(), ItemError>>> {
        let io: IOPCSyncIO = self.object.cast()?;
        let handles: Vec<_> = values.iter().map(|value| value.server_handle.0).collect();
        let variants: Vec<VARIANT> = values.iter().map(|value| value.value.clone()).collect();
        item_errors(values.len(), |errors| unsafe {
            io.Write(
                count(values.len())?,
                handles.as_ptr(),
                variants.as_ptr(),
                errors,
            )
        })
    }

    pub fn close(mut self, force: bool) -> Result<()> {
        let handle = self.server_handle.take();
        self.remove_on_drop = false;
        match handle {
            Some(handle) => unsafe { self.server.RemoveGroup(handle, force) },
            None => Ok(()),
        }
    }

    pub fn as_raw(&self) -> &IUnknown {
        &self.object
    }
}

fn validate_blob_lengths(specs: &[ItemSpec]) -> Result<()> {
    if specs
        .iter()
        .any(|spec| u32::try_from(spec.blob.len()).is_err())
    {
        Err(Error::from_hresult(
            windows::Win32::Foundation::E_INVALIDARG,
        ))
    } else {
        Ok(())
    }
}

fn blob_ptr(blob: &[u8]) -> *mut u8 {
    if blob.is_empty() {
        std::ptr::null_mut()
    } else {
        blob.as_ptr().cast_mut()
    }
}

impl Drop for DaGroup<'_> {
    fn drop(&mut self) {
        if self.remove_on_drop
            && let Some(handle) = self.server_handle.take()
        {
            let _ = unsafe { self.server.RemoveGroup(handle, false) };
        }
    }
}

fn item_errors(
    len: usize,
    call: impl FnOnce(*mut *mut HRESULT) -> Result<()>,
) -> Result<Vec<std::result::Result<(), ItemError>>> {
    let mut errors = CoTaskMemOut::<HRESULT>::new();
    let result = call(errors.as_mut_ptr());
    let errors = unsafe { errors.into_array(len, NoCleanup) }?;
    result?;
    Ok(errors
        .as_slice()
        .iter()
        .map(|error| {
            if error.is_err() {
                Err(ItemError { code: *error })
            } else {
                Ok(())
            }
        })
        .collect())
}

fn raw_handles(handles: &[ServerItemHandle]) -> Vec<u32> {
    handles.iter().map(|handle| handle.0).collect()
}

pub(crate) fn count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))
}

pub(crate) fn wide(value: &str) -> Result<WideCString> {
    WideCString::try_from(value)
        .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))
}
