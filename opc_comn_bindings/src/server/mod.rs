//! Rust service traits and COM adapters for OPC Common interfaces.

use std::sync::{Arc, Mutex};

use opc_classic_utils::server::{borrow_input, catch_ffi, initialize_output};
use opc_classic_utils::{CoTaskMemArrayBuilder, NoCleanup, OwnedPwstr};
use windows::Win32::Foundation::S_FALSE;
use windows_core::{Error, GUID, PCWSTR, PWSTR, Result};

use crate::{
    IOPCCommon, IOPCCommon_Impl, IOPCEnumGUID, IOPCEnumGUID_Impl, IOPCServerList2,
    IOPCServerList2_Impl, IOPCShutdown, IOPCShutdown_Impl,
};

pub trait CommonService: Send + Sync + 'static {
    fn set_locale(&self, locale: u32) -> Result<()>;
    fn locale(&self) -> Result<u32>;
    fn available_locales(&self) -> Result<Vec<u32>>;
    fn error_string(&self, error: windows_core::HRESULT) -> Result<String>;
    fn set_client_name(&self, name: String) -> Result<()>;
}

#[windows_core::implement(IOPCCommon)]
pub struct CommonServer {
    service: Arc<dyn CommonService>,
}

impl CommonServer {
    pub fn new(service: Arc<dyn CommonService>) -> Self {
        Self { service }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCCommon_Impl for CommonServer_Impl {
    fn SetLocaleID(&self, locale: u32) -> Result<()> {
        catch_ffi(|| self.service.set_locale(locale))
    }

    fn GetLocaleID(&self) -> Result<u32> {
        catch_ffi(|| self.service.locale())
    }

    fn QueryAvailableLocaleIDs(&self, count: *mut u32, values: *mut *mut u32) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(values)?;
            }
            let locales = self.service.available_locales()?;
            let count_value = u32::try_from(locales.len())
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_OUTOFMEMORY))?;
            let mut output = CoTaskMemArrayBuilder::new(locales.len(), NoCleanup)?;
            for locale in locales {
                output
                    .push(locale)
                    .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))?;
            }
            let output = output.finish()?;
            let (ptr, _) = output.into_raw_parts();
            unsafe {
                count.write(count_value);
                values.write(ptr);
            }
            Ok(())
        })
    }

    fn GetErrorString(&self, error: windows_core::HRESULT) -> Result<PWSTR> {
        catch_ffi(|| {
            let value = self.service.error_string(error)?;
            Ok(OwnedPwstr::new(value)?.into_raw())
        })
    }

    fn SetClientName(&self, name: &PCWSTR) -> Result<()> {
        catch_ffi(|| {
            let name = unsafe { name.to_string() }?;
            self.service.set_client_name(name)
        })
    }
}

#[windows_core::implement(IOPCEnumGUID)]
struct GuidEnumeratorServer {
    values: Arc<Vec<GUID>>,
    position: Mutex<usize>,
}

impl GuidEnumeratorServer {
    fn new(values: Vec<GUID>) -> Self {
        Self {
            values: Arc::new(values),
            position: Mutex::new(0),
        }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCEnumGUID_Impl for GuidEnumeratorServer_Impl {
    fn Next(&self, count: u32, values: *mut GUID, fetched: *mut u32) -> Result<()> {
        catch_ffi(|| {
            unsafe { initialize_output(fetched)? };
            if count != 0 && values.is_null() {
                return Err(Error::from_hresult(windows::Win32::Foundation::E_POINTER));
            }
            let mut position = self
                .position
                .lock()
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))?;
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
                Err(Error::from_hresult(S_FALSE))
            } else {
                Ok(())
            }
        })
    }

    fn Skip(&self, count: u32) -> Result<()> {
        catch_ffi(|| {
            let mut position = self
                .position
                .lock()
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))?;
            let available = self.values.len().saturating_sub(*position);
            let skipped = available.min(count as usize);
            *position += skipped;
            if skipped < count as usize {
                Err(Error::from_hresult(S_FALSE))
            } else {
                Ok(())
            }
        })
    }

    fn Reset(&self) -> Result<()> {
        catch_ffi(|| {
            *self
                .position
                .lock()
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))? = 0;
            Ok(())
        })
    }

    fn Clone(&self) -> Result<IOPCEnumGUID> {
        catch_ffi(|| {
            let position = *self
                .position
                .lock()
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))?;
            let clone = GuidEnumeratorServer {
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
    fn enum_classes(&self, implemented: &[GUID], required: &[GUID]) -> Result<Vec<GUID>>;
    fn class_details(&self, class_id: &GUID) -> Result<ClassDetails>;
    fn class_id_from_prog_id(&self, prog_id: &str) -> Result<GUID>;
}

#[windows_core::implement(IOPCServerList2)]
pub struct ServerListServer {
    service: Arc<dyn ServerListService>,
}

impl ServerListServer {
    pub fn new(service: Arc<dyn ServerListService>) -> Self {
        Self { service }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCServerList2_Impl for ServerListServer_Impl {
    fn EnumClassesOfCategories(
        &self,
        implemented_count: u32,
        implemented: *const GUID,
        required_count: u32,
        required: *const GUID,
    ) -> Result<IOPCEnumGUID> {
        catch_ffi(|| {
            let implemented = unsafe { borrow_input(implemented, implemented_count)? };
            let required = unsafe { borrow_input(required, required_count)? };
            let values = self.service.enum_classes(implemented, required)?;
            Ok(GuidEnumeratorServer::new(values).into())
        })
    }

    fn GetClassDetails(
        &self,
        class_id: *const GUID,
        prog_id: *mut PWSTR,
        user_type: *mut PWSTR,
        version_independent: *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            if class_id.is_null() {
                return Err(Error::from_hresult(windows::Win32::Foundation::E_POINTER));
            }
            unsafe {
                initialize_output(prog_id)?;
                initialize_output(user_type)?;
                initialize_output(version_independent)?;
            }
            let details = self.service.class_details(unsafe { &*class_id })?;
            let prog_id_value = OwnedPwstr::new(details.prog_id)?;
            let user_type_value = OwnedPwstr::new(details.user_type)?;
            let version_value = OwnedPwstr::new(details.version_independent_prog_id)?;
            unsafe {
                prog_id.write(prog_id_value.into_raw());
                user_type.write(user_type_value.into_raw());
                version_independent.write(version_value.into_raw());
            }
            Ok(())
        })
    }

    fn CLSIDFromProgID(&self, prog_id: &PCWSTR) -> Result<GUID> {
        catch_ffi(|| {
            self.service
                .class_id_from_prog_id(&unsafe { prog_id.to_string() }?)
        })
    }
}

pub trait ShutdownHandler: Send + Sync + 'static {
    fn shutdown_requested(&self, reason: String) -> Result<()>;
}

#[windows_core::implement(IOPCShutdown)]
pub struct ShutdownServer {
    handler: Arc<dyn ShutdownHandler>,
}

impl ShutdownServer {
    pub fn new(handler: Arc<dyn ShutdownHandler>) -> Self {
        Self { handler }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCShutdown_Impl for ShutdownServer_Impl {
    fn ShutdownRequest(&self, reason: &PCWSTR) -> Result<()> {
        catch_ffi(|| {
            self.handler
                .shutdown_requested(unsafe { reason.to_string() }?)
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::client::CommonClient;

    struct TestCommon;

    impl CommonService for TestCommon {
        fn set_locale(&self, _locale: u32) -> Result<()> {
            Ok(())
        }

        fn locale(&self) -> Result<u32> {
            Ok(1_033)
        }

        fn available_locales(&self) -> Result<Vec<u32>> {
            Ok(vec![1_033, 2_052])
        }

        fn error_string(&self, _error: windows_core::HRESULT) -> Result<String> {
            Ok("test error".into())
        }

        fn set_client_name(&self, _name: String) -> Result<()> {
            Ok(())
        }
    }

    #[test]
    fn common_adapter_round_trips_task_memory() {
        let raw: IOPCCommon = CommonServer::new(Arc::new(TestCommon)).into();
        let client = CommonClient::new(raw);
        assert_eq!(client.locale().unwrap(), 1_033);
        assert_eq!(client.available_locales().unwrap(), [1_033, 2_052]);
        assert_eq!(
            client.error_string(windows_core::HRESULT(1)).unwrap(),
            "test error"
        );
    }
}
