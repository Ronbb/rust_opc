//! Rust service traits and COM adapters for OPC Common interfaces.

use std::sync::{Arc, Mutex};

use opc_classic_types::{ComObject, Error, ErrorCode, Guid, Result};
use opc_classic_utils::server::{borrow_input, catch_abi, initialize_output};
use opc_classic_utils::{CoTaskMemArrayBuilder, NoCleanup, OwnedPwstr};
use windows::Win32::Foundation::{E_UNEXPECTED, S_FALSE};
use windows_core::{Error as AbiError, GUID, HRESULT, PCWSTR, PWSTR, Result as AbiResult};

use crate::abi::{from_abi_guid, object_from_interface, to_abi_error, to_abi_guid};
use crate::{
    IOPCCommon, IOPCCommon_Impl, IOPCEnumGUID, IOPCEnumGUID_Impl, IOPCServerList2,
    IOPCServerList2_Impl, IOPCShutdown, IOPCShutdown_Impl,
};

fn abi_boundary<T>(f: impl FnOnce() -> AbiResult<T>) -> AbiResult<T> {
    catch_abi(f, AbiError::from_hresult(E_UNEXPECTED))
}

pub trait CommonService: Send + Sync + 'static {
    fn set_locale(&self, locale: u32) -> Result<()>;
    fn locale(&self) -> Result<u32>;
    fn available_locales(&self) -> Result<Vec<u32>>;
    fn error_string(&self, error: ErrorCode) -> Result<String>;
    fn set_client_name(&self, name: String) -> Result<()>;
}

#[windows_core::implement(IOPCCommon)]
struct CommonAdapter {
    service: Arc<dyn CommonService>,
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCCommon_Impl for CommonAdapter_Impl {
    fn SetLocaleID(&self, locale: u32) -> AbiResult<()> {
        abi_boundary(|| self.service.set_locale(locale).map_err(to_abi_error))
    }

    fn GetLocaleID(&self) -> AbiResult<u32> {
        abi_boundary(|| self.service.locale().map_err(to_abi_error))
    }

    fn QueryAvailableLocaleIDs(&self, count: *mut u32, values: *mut *mut u32) -> AbiResult<()> {
        abi_boundary(|| {
            unsafe {
                initialize_output(count).map_err(to_abi_error)?;
                initialize_output(values).map_err(to_abi_error)?;
            }
            let locales = self.service.available_locales().map_err(to_abi_error)?;
            let count_value = u32::try_from(locales.len()).map_err(|_| {
                to_abi_error(Error::out_of_memory(
                    "locale array is too large for the OPC ABI",
                ))
            })?;
            let mut output =
                CoTaskMemArrayBuilder::new(locales.len(), NoCleanup).map_err(to_abi_error)?;
            for locale in locales {
                output
                    .push(locale)
                    .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))?;
            }
            let output = output.finish().map_err(to_abi_error)?;
            let (ptr, _) = output.into_raw_parts();
            unsafe {
                count.write(count_value);
                values.write(ptr);
            }
            Ok(())
        })
    }

    fn GetErrorString(&self, error: HRESULT) -> AbiResult<PWSTR> {
        abi_boundary(|| {
            let value = self
                .service
                .error_string(ErrorCode::from_raw(error.0))
                .map_err(to_abi_error)?;
            let value = OwnedPwstr::new(value).map_err(to_abi_error)?;
            Ok(PWSTR(value.into_raw()))
        })
    }

    fn SetClientName(&self, name: &PCWSTR) -> AbiResult<()> {
        abi_boundary(|| {
            if name.is_null() {
                return Err(to_abi_error(Error::null_pointer("null client name")));
            }
            let name = unsafe { name.to_string() }?;
            self.service.set_client_name(name).map_err(to_abi_error)
        })
    }
}

/// An OPC Common server backed by a Rust [`CommonService`].
#[derive(Clone)]
pub struct CommonServer {
    object: ComObject,
}

impl CommonServer {
    pub fn new(service: Arc<dyn CommonService>) -> Self {
        let interface: IOPCCommon = CommonAdapter { service }.into();
        Self {
            object: object_from_interface(&interface),
        }
    }

    pub fn object(&self) -> ComObject {
        self.object.clone()
    }
}

#[windows_core::implement(IOPCEnumGUID)]
struct GuidEnumeratorAdapter {
    values: Arc<Vec<GUID>>,
    position: Mutex<usize>,
}

impl GuidEnumeratorAdapter {
    fn new(values: Vec<GUID>) -> Self {
        Self {
            values: Arc::new(values),
            position: Mutex::new(0),
        }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCEnumGUID_Impl for GuidEnumeratorAdapter_Impl {
    fn Next(&self, count: u32, values: *mut GUID, fetched: *mut u32) -> AbiResult<()> {
        abi_boundary(|| {
            unsafe { initialize_output(fetched).map_err(to_abi_error)? };
            if count != 0 && values.is_null() {
                return Err(to_abi_error(Error::null_pointer(
                    "null GUID enumeration output",
                )));
            }
            let mut position = self
                .position
                .lock()
                .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))?;
            let available = self.values.len().saturating_sub(*position);
            let returned = available.min(count as usize);
            if returned != 0 {
                unsafe {
                    std::ptr::copy_nonoverlapping(
                        self.values.as_ptr().add(*position),
                        values,
                        returned,
                    );
                }
            }
            *position += returned;
            unsafe { fetched.write(returned as u32) };
            if returned < count as usize {
                Err(AbiError::from_hresult(S_FALSE))
            } else {
                Ok(())
            }
        })
    }

    fn Skip(&self, count: u32) -> AbiResult<()> {
        abi_boundary(|| {
            let mut position = self
                .position
                .lock()
                .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))?;
            let available = self.values.len().saturating_sub(*position);
            let skipped = available.min(count as usize);
            *position += skipped;
            if skipped < count as usize {
                Err(AbiError::from_hresult(S_FALSE))
            } else {
                Ok(())
            }
        })
    }

    fn Reset(&self) -> AbiResult<()> {
        abi_boundary(|| {
            *self
                .position
                .lock()
                .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))? = 0;
            Ok(())
        })
    }

    fn Clone(&self) -> AbiResult<IOPCEnumGUID> {
        abi_boundary(|| {
            let position = *self
                .position
                .lock()
                .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))?;
            let clone = GuidEnumeratorAdapter {
                values: self.values.clone(),
                position: Mutex::new(position),
            };
            Ok(clone.into())
        })
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ClassDetails {
    pub prog_id: String,
    pub user_type: String,
    pub version_independent_prog_id: String,
}

pub trait ServerListService: Send + Sync + 'static {
    fn enum_classes(&self, implemented: &[Guid], required: &[Guid]) -> Result<Vec<Guid>>;
    fn class_details(&self, class_id: &Guid) -> Result<ClassDetails>;
    fn class_id_from_prog_id(&self, prog_id: &str) -> Result<Guid>;
}

#[windows_core::implement(IOPCServerList2)]
struct ServerListAdapter {
    service: Arc<dyn ServerListService>,
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCServerList2_Impl for ServerListAdapter_Impl {
    fn EnumClassesOfCategories(
        &self,
        implemented_count: u32,
        implemented: *const GUID,
        required_count: u32,
        required: *const GUID,
    ) -> AbiResult<IOPCEnumGUID> {
        abi_boundary(|| {
            let implemented =
                unsafe { borrow_input(implemented, implemented_count) }.map_err(to_abi_error)?;
            let required =
                unsafe { borrow_input(required, required_count) }.map_err(to_abi_error)?;
            let implemented: Vec<_> = implemented.iter().copied().map(from_abi_guid).collect();
            let required: Vec<_> = required.iter().copied().map(from_abi_guid).collect();
            let values = self
                .service
                .enum_classes(&implemented, &required)
                .map_err(to_abi_error)?;
            let values = values.into_iter().map(to_abi_guid).collect();
            Ok(GuidEnumeratorAdapter::new(values).into())
        })
    }

    fn GetClassDetails(
        &self,
        class_id: *const GUID,
        prog_id: *mut PWSTR,
        user_type: *mut PWSTR,
        version_independent: *mut PWSTR,
    ) -> AbiResult<()> {
        abi_boundary(|| {
            unsafe {
                initialize_output(prog_id).map_err(to_abi_error)?;
                initialize_output(user_type).map_err(to_abi_error)?;
                initialize_output(version_independent).map_err(to_abi_error)?;
            }
            let class_id = unsafe { class_id.as_ref() }
                .copied()
                .ok_or_else(|| to_abi_error(Error::null_pointer("null class ID")))?;
            let details = self
                .service
                .class_details(&from_abi_guid(class_id))
                .map_err(to_abi_error)?;
            let prog_id_value = OwnedPwstr::new(details.prog_id).map_err(to_abi_error)?;
            let user_type_value = OwnedPwstr::new(details.user_type).map_err(to_abi_error)?;
            let version_value =
                OwnedPwstr::new(details.version_independent_prog_id).map_err(to_abi_error)?;
            unsafe {
                prog_id.write(PWSTR(prog_id_value.into_raw()));
                user_type.write(PWSTR(user_type_value.into_raw()));
                version_independent.write(PWSTR(version_value.into_raw()));
            }
            Ok(())
        })
    }

    fn CLSIDFromProgID(&self, prog_id: &PCWSTR) -> AbiResult<GUID> {
        abi_boundary(|| {
            if prog_id.is_null() {
                return Err(to_abi_error(Error::null_pointer("null ProgID")));
            }
            let prog_id = unsafe { prog_id.to_string() }?;
            self.service
                .class_id_from_prog_id(&prog_id)
                .map(to_abi_guid)
                .map_err(to_abi_error)
        })
    }
}

/// An OPC Server List server backed by a Rust [`ServerListService`].
#[derive(Clone)]
pub struct ServerListServer {
    object: ComObject,
}

impl ServerListServer {
    pub fn new(service: Arc<dyn ServerListService>) -> Self {
        let interface: IOPCServerList2 = ServerListAdapter { service }.into();
        Self {
            object: object_from_interface(&interface),
        }
    }

    pub fn object(&self) -> ComObject {
        self.object.clone()
    }
}

pub trait ShutdownHandler: Send + Sync + 'static {
    fn shutdown_requested(&self, reason: String) -> Result<()>;
}

#[windows_core::implement(IOPCShutdown)]
struct ShutdownAdapter {
    handler: Arc<dyn ShutdownHandler>,
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCShutdown_Impl for ShutdownAdapter_Impl {
    fn ShutdownRequest(&self, reason: &PCWSTR) -> AbiResult<()> {
        abi_boundary(|| {
            if reason.is_null() {
                return Err(to_abi_error(Error::null_pointer("null shutdown reason")));
            }
            let reason = unsafe { reason.to_string() }?;
            self.handler
                .shutdown_requested(reason)
                .map_err(to_abi_error)
        })
    }
}

/// A shutdown callback object backed by a Rust [`ShutdownHandler`].
#[derive(Clone)]
pub struct ShutdownServer {
    object: ComObject,
}

impl ShutdownServer {
    pub fn new(handler: Arc<dyn ShutdownHandler>) -> Self {
        let interface: IOPCShutdown = ShutdownAdapter { handler }.into();
        Self {
            object: object_from_interface(&interface),
        }
    }

    pub fn object(&self) -> ComObject {
        self.object.clone()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicBool, Ordering};

    use crate::client::{CommonClient, ServerListClient, ShutdownClient};

    const CLASS_ID: Guid = Guid::new(
        0x1234_5678,
        0x1234,
        0x5678,
        [0x90, 0xab, 0xcd, 0xef, 1, 2, 3, 4],
    );

    struct TestCommon;

    impl CommonService for TestCommon {
        fn set_locale(&self, locale: u32) -> Result<()> {
            if locale == 0 {
                Err(Error::invalid_argument("invalid test locale"))
            } else {
                Ok(())
            }
        }

        fn locale(&self) -> Result<u32> {
            Ok(1_033)
        }

        fn available_locales(&self) -> Result<Vec<u32>> {
            Ok(vec![1_033, 2_052])
        }

        fn error_string(&self, _error: ErrorCode) -> Result<String> {
            Ok("test error".into())
        }

        fn set_client_name(&self, _name: String) -> Result<()> {
            Ok(())
        }
    }

    #[test]
    fn common_adapter_round_trips_task_memory() {
        let server = CommonServer::new(Arc::new(TestCommon));
        let client = CommonClient::from_object(&server.object()).unwrap();
        assert_eq!(client.locale().unwrap(), 1_033);
        assert_eq!(client.available_locales().unwrap(), [1_033, 2_052]);
        assert_eq!(
            client.error_string(ErrorCode::PARTIAL_SUCCESS).unwrap(),
            "test error"
        );
        assert_eq!(
            client.set_locale(0).unwrap_err().code(),
            ErrorCode::INVALID_ARGUMENT
        );
    }

    struct TestServerList;

    impl ServerListService for TestServerList {
        fn enum_classes(&self, implemented: &[Guid], required: &[Guid]) -> Result<Vec<Guid>> {
            assert_eq!(implemented, [CLASS_ID]);
            assert!(required.is_empty());
            Ok(vec![CLASS_ID])
        }

        fn class_details(&self, class_id: &Guid) -> Result<ClassDetails> {
            assert_eq!(*class_id, CLASS_ID);
            Ok(ClassDetails {
                prog_id: "Example.OPC.1".into(),
                user_type: "Example OPC Server".into(),
                version_independent_prog_id: "Example.OPC".into(),
            })
        }

        fn class_id_from_prog_id(&self, prog_id: &str) -> Result<Guid> {
            assert_eq!(prog_id, "Example.OPC.1");
            Ok(CLASS_ID)
        }
    }

    #[test]
    fn server_list_adapter_converts_owned_guids_and_strings() {
        let server = ServerListServer::new(Arc::new(TestServerList));
        let client = ServerListClient::from_object(&server.object()).unwrap();

        let enumerator = client.enum_classes(&[CLASS_ID], &[]).unwrap();
        assert_eq!(enumerator.next_batch(2).unwrap(), [CLASS_ID]);
        assert!(enumerator.next_batch(1).unwrap().is_empty());

        let details = client.class_details(&CLASS_ID).unwrap();
        assert_eq!(details.prog_id, "Example.OPC.1");
        assert_eq!(details.user_type, "Example OPC Server");
        assert_eq!(details.version_independent_prog_id, "Example.OPC");
        assert_eq!(
            client.class_id_from_prog_id("Example.OPC.1").unwrap(),
            CLASS_ID
        );
    }

    struct TestShutdown {
        called: Arc<AtomicBool>,
    }

    impl ShutdownHandler for TestShutdown {
        fn shutdown_requested(&self, reason: String) -> Result<()> {
            assert_eq!(reason, "maintenance");
            self.called.store(true, Ordering::Release);
            Ok(())
        }
    }

    #[test]
    fn shutdown_adapter_accepts_rust_strings() {
        let called = Arc::new(AtomicBool::new(false));
        let server = ShutdownServer::new(Arc::new(TestShutdown {
            called: called.clone(),
        }));
        let client = ShutdownClient::from_object(&server.object()).unwrap();

        client.request("maintenance").unwrap();
        assert!(called.load(Ordering::Acquire));
        assert!(client.request("bad\0reason").is_err());
    }
}
