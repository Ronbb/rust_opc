//! Safe OPC Data Access client facade.

mod group;
mod properties;
mod types;

use std::marker::PhantomData;

use opc_classic_types::{ClassContext, ComObject, Error, ErrorCode, Guid, Result};
use opc_classic_utils::{Cleanup, CoTaskMemObjectOut, CoTaskMemOut, ComApartment};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::System::Com::{CLSIDFromProgID, CoCreateInstance, CoTaskMemFree};
use windows_core::{IUnknown, Interface, PCWSTR};

use crate::abi::{
    class_context_to_abi, error_code_to_abi, from_abi_error, guid_to_abi, interface_from_object,
    object_from_interface, timestamp_from_abi,
};
use crate::{IOPCItemMgt, IOPCServer, tagOPCSERVERSTATUS};

pub use group::DaGroup;
pub use properties::ItemPropertiesClient;
pub use types::*;

struct ServerStatusCleanup;

// SAFETY: OPCSERVERSTATUS owns its vendor string and the outer structure.
unsafe impl Cleanup<tagOPCSERVERSTATUS> for ServerStatusCleanup {
    unsafe fn cleanup(&mut self, ptr: *mut tagOPCSERVERSTATUS, initialized: usize) {
        for index in 0..initialized {
            let status = unsafe { &mut *ptr.add(index) };
            if !status.szVendorInfo.is_null() {
                unsafe { CoTaskMemFree(Some(status.szVendorInfo.0.cast())) };
                status.szVendorInfo = windows_core::PWSTR::null();
            }
        }
    }
}

#[derive(Clone)]
pub struct DaClient<'apartment> {
    inner: IOPCServer,
    _apartment: PhantomData<&'apartment ComApartment>,
}

impl<'apartment> DaClient<'apartment> {
    pub fn connect_prog_id(apartment: &'apartment ComApartment, prog_id: &str) -> Result<Self> {
        let prog_id = group::wide(prog_id)?;
        let class_id =
            unsafe { CLSIDFromProgID(PCWSTR(prog_id.as_ptr())) }.map_err(from_abi_error)?;
        let class_id = Guid::new(
            class_id.data1,
            class_id.data2,
            class_id.data3,
            class_id.data4,
        );
        Self::connect_class_id(apartment, &class_id, ClassContext::ALL)
    }

    pub fn connect_class_id(
        apartment: &'apartment ComApartment,
        class_id: &Guid,
        context: ClassContext,
    ) -> Result<Self> {
        let class_id = guid_to_abi(class_id);
        let context = class_context_to_abi(context);
        let inner = unsafe { CoCreateInstance(&class_id, None::<&IUnknown>, context) }
            .map_err(from_abi_error)?;
        Ok(Self::from_interface(apartment, inner))
    }

    pub fn from_object(apartment: &'apartment ComApartment, object: &ComObject) -> Result<Self> {
        let inner = interface_from_object(object)?;
        Ok(Self::from_interface(apartment, inner))
    }

    pub(crate) fn from_interface(apartment: &'apartment ComApartment, inner: IOPCServer) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn common(&self) -> Result<CommonClient> {
        CommonClient::from_object(&self.object())
    }

    pub fn properties(&self) -> Result<ItemPropertiesClient> {
        ItemPropertiesClient::from_object(&self.object())
    }

    pub fn add_group(&self, options: GroupOptions) -> Result<DaGroup<'apartment>> {
        let name = group::wide(&options.name)?;
        let time_bias = options.time_bias.as_ref().map_or(std::ptr::null(), |v| v);
        let deadband = options
            .percent_deadband
            .as_ref()
            .map_or(std::ptr::null(), |v| v);
        let mut server_handle = 0;
        let mut revised_update_rate = 0;
        let mut object = None;
        unsafe {
            self.inner.AddGroup(
                PCWSTR(name.as_ptr()),
                options.active,
                options.requested_update_rate,
                options.client_handle,
                time_bias,
                deadband,
                options.locale,
                &mut server_handle,
                &mut revised_update_rate,
                &IOPCItemMgt::IID,
                &mut object,
            )
        }
        .map_err(from_abi_error)?;
        let object = match object {
            Some(object) => object,
            None => {
                if server_handle != 0
                    && let Err(cleanup_error) =
                        unsafe { self.inner.RemoveGroup(server_handle, true) }
                {
                    let cleanup_error = from_abi_error(cleanup_error);
                    return Err(Error::unexpected(format!(
                        "AddGroup returned no group object and cleanup of server group {server_handle} failed: {cleanup_error}"
                    )));
                }
                return Err(Error::unexpected("AddGroup returned no group object"));
            }
        };
        Ok(DaGroup::new(
            self.inner.clone(),
            object,
            server_handle,
            true,
            revised_update_rate,
        ))
    }

    pub fn status(&self) -> Result<ServerStatus> {
        let mut status = CoTaskMemObjectOut::new(ServerStatusCleanup);
        let call = unsafe {
            (Interface::vtable(&self.inner).GetStatus)(
                Interface::as_raw(&self.inner),
                status.as_mut_ptr(),
            )
        };
        let status = unsafe { status.into_object() };
        call.ok().map_err(from_abi_error)?;
        let status = status?;
        let raw = status.as_ref();
        let vendor_info = if raw.szVendorInfo.is_null() {
            String::new()
        } else {
            unsafe { raw.szVendorInfo.to_string() }
                .map_err(|_| Error::invalid_argument("invalid UTF-16 vendor string"))?
        };
        Ok(ServerStatus {
            start_time: timestamp_from_abi(raw.ftStartTime),
            current_time: timestamp_from_abi(raw.ftCurrentTime),
            last_update_time: timestamp_from_abi(raw.ftLastUpdateTime),
            state: ServerState::from_raw(raw.dwServerState.0),
            group_count: raw.dwGroupCount,
            bandwidth: raw.dwBandWidth,
            version: (raw.wMajorVersion, raw.wMinorVersion, raw.wBuildNumber),
            vendor_info,
        })
    }

    pub fn error_string(&self, error: ErrorCode, locale: u32) -> Result<String> {
        let mut value = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            (Interface::vtable(&self.inner).GetErrorString)(
                Interface::as_raw(&self.inner),
                error_code_to_abi(error),
                locale,
                value.as_mut_ptr().cast(),
            )
        };
        let value = unsafe { value.into_pwstr() };
        call.ok().map_err(from_abi_error)?;
        Ok(value.to_string_lossy())
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Arc, Mutex};

    use crate::{IOPCServer_Impl, tagOPCENUMSCOPE};
    use windows::Win32::Foundation::E_NOTIMPL;
    use windows_core::{Error as AbiError, GUID, HRESULT, OutRef, PWSTR, Result as AbiResult};

    #[windows_core::implement(IOPCServer)]
    struct MissingGroupObjectServer {
        removed: Arc<Mutex<Vec<(u32, bool)>>>,
    }

    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    impl IOPCServer_Impl for MissingGroupObjectServer_Impl {
        fn AddGroup(
            &self,
            _name: &PCWSTR,
            _active: windows_core::BOOL,
            requested_update_rate: u32,
            _client_handle: u32,
            _time_bias: *const i32,
            _percent_deadband: *const f32,
            _locale: u32,
            server_handle: *mut u32,
            revised_update_rate: *mut u32,
            _iid: *const GUID,
            output: OutRef<IUnknown>,
        ) -> AbiResult<()> {
            unsafe {
                server_handle.write(73);
                revised_update_rate.write(requested_update_rate);
            }
            output.write(None)
        }

        fn GetErrorString(&self, _error: HRESULT, _locale: u32) -> AbiResult<PWSTR> {
            Err(AbiError::from_hresult(E_NOTIMPL))
        }

        fn GetGroupByName(&self, _name: &PCWSTR, _iid: *const GUID) -> AbiResult<IUnknown> {
            Err(AbiError::from_hresult(E_NOTIMPL))
        }

        fn GetStatus(&self) -> AbiResult<*mut tagOPCSERVERSTATUS> {
            Err(AbiError::from_hresult(E_NOTIMPL))
        }

        fn RemoveGroup(&self, server_handle: u32, force: windows_core::BOOL) -> AbiResult<()> {
            self.removed
                .lock()
                .expect("test removal log should not be poisoned")
                .push((server_handle, force.as_bool()));
            Ok(())
        }

        fn CreateGroupEnumerator(
            &self,
            _scope: tagOPCENUMSCOPE,
            _iid: *const GUID,
        ) -> AbiResult<IUnknown> {
            Err(AbiError::from_hresult(E_NOTIMPL))
        }
    }

    #[test]
    fn add_group_removes_server_group_when_object_is_missing() {
        let apartment = ComApartment::mta().unwrap();
        let removed = Arc::new(Mutex::new(Vec::new()));
        let server: IOPCServer = MissingGroupObjectServer {
            removed: removed.clone(),
        }
        .into();
        let client = DaClient::from_interface(&apartment, server);

        let error = client
            .add_group(GroupOptions::new("missing object"))
            .err()
            .expect("missing group object should be rejected");

        assert_eq!(error.code(), ErrorCode::UNEXPECTED);
        assert_eq!(
            *removed
                .lock()
                .expect("test removal log should not be poisoned"),
            [(73, true)]
        );
    }
}
