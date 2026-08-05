//! Rust service contracts for OPC Alarms & Events servers.

use std::sync::Arc;

use opc_classic_utils::server::{borrow_input, catch_ffi, initialize_output};
use opc_classic_utils::{Cleanup, CoTaskMemArrayBuilder, FreePwstrElements, NoCleanup, OwnedPwstr};
use windows::Win32::Foundation::{E_NOTIMPL, E_UNEXPECTED, FILETIME};
use windows::Win32::System::Com::CoTaskMemFree;
use windows_core::{Error, GUID, HRESULT, IUnknown, OutRef, PCWSTR, PWSTR, Result};

use crate::{IOPCEventServer, IOPCEventServer_Impl};

#[derive(Clone, Debug)]
pub struct AeStatus {
    pub start_time: FILETIME,
    pub current_time: FILETIME,
    pub last_update_time: FILETIME,
    pub state: crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0003,
    pub version: (u16, u16, u16),
    pub vendor_info: String,
}

pub trait AeService: Send + Sync + 'static {
    fn status(&self) -> Result<AeStatus>;
    fn available_filters(&self) -> Result<u32>;
    fn event_categories(&self, event_type: u32) -> Result<Vec<(u32, String)>>;
    fn condition_names(&self, event_category: u32) -> Result<Vec<String>>;
    fn subcondition_names(&self, condition: &str) -> Result<Vec<String>>;
    fn source_conditions(&self, source: &str) -> Result<Vec<String>>;
    fn event_attributes(&self, event_category: u32) -> Result<Vec<(u32, String, u16)>>;
    fn enable_area(&self, names: &[String], enabled: bool) -> Result<()>;
    fn enable_source(&self, names: &[String], enabled: bool) -> Result<()>;
}

/// Shared service handle used by an application-specific COM adapter.
#[derive(Clone)]
pub struct AeServiceHandle(pub Arc<dyn AeService>);

impl AeServiceHandle {
    pub fn new(service: Arc<dyn AeService>) -> Self {
        Self(service)
    }
}

/// Per-item/event HRESULTs are represented explicitly instead of being folded
/// into a transport-level error.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EventError(pub HRESULT);

#[windows_core::implement(IOPCEventServer)]
pub struct AeServerAdapter {
    service: Arc<dyn AeService>,
}

impl AeServerAdapter {
    pub fn new(service: Arc<dyn AeService>) -> Self {
        Self { service }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCEventServer_Impl for AeServerAdapter_Impl {
    fn GetStatus(&self) -> Result<*mut crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005> {
        catch_ffi(|| {
            let status = self.service.status()?;
            let vendor = OwnedPwstr::new(status.vendor_info)?;
            let mut output = CoTaskMemArrayBuilder::new(1, StatusCleanup)?;
            output
                .push(crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005 {
                    ftStartTime: status.start_time,
                    ftCurrentTime: status.current_time,
                    ftLastUpdateTime: status.last_update_time,
                    dwServerState: status.state,
                    wMajorVersion: status.version.0,
                    wMinorVersion: status.version.1,
                    wBuildNumber: status.version.2,
                    wReserved: 0,
                    szVendorInfo: vendor.into_raw(),
                })
                .map_err(|_| Error::from_hresult(E_UNEXPECTED))?;
            Ok(output.finish()?.into_raw_parts().0)
        })
    }

    fn CreateEventSubscription(
        &self,
        _active: windows_core::BOOL,
        _buffer_time: u32,
        _max_size: u32,
        _client_subscription: u32,
        _iid: *const GUID,
        _output: OutRef<IUnknown>,
        _revised_buffer_time: *mut u32,
        _revised_max_size: *mut u32,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn QueryAvailableFilters(&self) -> Result<u32> {
        catch_ffi(|| self.service.available_filters())
    }

    fn QueryEventCategories(
        &self,
        event_type: u32,
        count: *mut u32,
        ids: *mut *mut u32,
        descriptions: *mut *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(ids)?;
                initialize_output(descriptions)?;
            }
            let values = self.service.event_categories(event_type)?;
            let count_value = checked_count(values.len())?;
            let mut id_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            let mut description_values =
                CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            for (id, description) in values {
                id_values
                    .push(id)
                    .map_err(|_| Error::from_hresult(E_UNEXPECTED))?;
                push_pwstr(&mut description_values, description)?;
            }
            let id_values = id_values.finish()?.into_raw_parts().0;
            let description_values = description_values.finish()?.into_raw_parts().0;
            unsafe {
                count.write(count_value);
                ids.write(id_values);
                descriptions.write(description_values);
            }
            Ok(())
        })
    }

    fn QueryConditionNames(
        &self,
        category: u32,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| write_strings(self.service.condition_names(category)?, count, names))
    }

    fn QuerySubConditionNames(
        &self,
        condition: &PCWSTR,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            let condition = unsafe { condition.to_string() }?;
            write_strings(self.service.subcondition_names(&condition)?, count, names)
        })
    }

    fn QuerySourceConditions(
        &self,
        source: &PCWSTR,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            let source = unsafe { source.to_string() }?;
            write_strings(self.service.source_conditions(&source)?, count, names)
        })
    }

    fn QueryEventAttributes(
        &self,
        category: u32,
        count: *mut u32,
        ids: *mut *mut u32,
        descriptions: *mut *mut PWSTR,
        data_types: *mut *mut u16,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(ids)?;
                initialize_output(descriptions)?;
                initialize_output(data_types)?;
            }
            let values = self.service.event_attributes(category)?;
            let count_value = checked_count(values.len())?;
            let mut id_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            let mut description_values =
                CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            let mut type_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            for (id, description, data_type) in values {
                id_values
                    .push(id)
                    .map_err(|_| Error::from_hresult(E_UNEXPECTED))?;
                push_pwstr(&mut description_values, description)?;
                type_values
                    .push(data_type)
                    .map_err(|_| Error::from_hresult(E_UNEXPECTED))?;
            }
            unsafe {
                count.write(count_value);
                ids.write(id_values.finish()?.into_raw_parts().0);
                descriptions.write(description_values.finish()?.into_raw_parts().0);
                data_types.write(type_values.finish()?.into_raw_parts().0);
            }
            Ok(())
        })
    }

    fn TranslateToItemIDs(
        &self,
        _source: &PCWSTR,
        _event_category: u32,
        _condition_name: &PCWSTR,
        _subcondition_name: &PCWSTR,
        _count: u32,
        _attribute_ids: *const u32,
        _item_ids: *mut *mut PWSTR,
        _node_names: *mut *mut PWSTR,
        _class_ids: *mut *mut GUID,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn GetConditionState(
        &self,
        _source: &PCWSTR,
        _condition_name: &PCWSTR,
        _count: u32,
        _attribute_ids: *const u32,
    ) -> Result<*mut crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0006> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn EnableConditionByArea(&self, count: u32, names: *const PCWSTR) -> Result<()> {
        catch_ffi(|| {
            self.service
                .enable_area(&decode_strings(count, names)?, true)
        })
    }

    fn EnableConditionBySource(&self, count: u32, names: *const PCWSTR) -> Result<()> {
        catch_ffi(|| {
            self.service
                .enable_source(&decode_strings(count, names)?, true)
        })
    }

    fn DisableConditionByArea(&self, count: u32, names: *const PCWSTR) -> Result<()> {
        catch_ffi(|| {
            self.service
                .enable_area(&decode_strings(count, names)?, false)
        })
    }

    fn DisableConditionBySource(&self, count: u32, names: *const PCWSTR) -> Result<()> {
        catch_ffi(|| {
            self.service
                .enable_source(&decode_strings(count, names)?, false)
        })
    }

    fn AckCondition(
        &self,
        _count: u32,
        _acknowledger: &PCWSTR,
        _comment: &PCWSTR,
        _sources: *const PCWSTR,
        _condition_names: *const PCWSTR,
        _active_times: *const FILETIME,
        _cookies: *const u32,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn CreateAreaBrowser(&self, _iid: *const GUID) -> Result<IUnknown> {
        Err(Error::from_hresult(E_NOTIMPL))
    }
}

struct StatusCleanup;

// SAFETY: The AE status owns one task-allocated vendor string.
unsafe impl Cleanup<crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005> for StatusCleanup {
    unsafe fn cleanup(
        &mut self,
        ptr: *mut crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005,
        initialized: usize,
    ) {
        for index in 0..initialized {
            let value = unsafe { &mut *ptr.add(index) };
            if !value.szVendorInfo.is_null() {
                unsafe { CoTaskMemFree(Some(value.szVendorInfo.0.cast())) };
                value.szVendorInfo = PWSTR::null();
            }
        }
    }
}

fn write_strings(values: Vec<String>, count: *mut u32, output: *mut *mut PWSTR) -> Result<()> {
    unsafe {
        initialize_output(count)?;
        initialize_output(output)?;
    }
    let count_value = checked_count(values.len())?;
    let mut strings = CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
    for value in values {
        push_pwstr(&mut strings, value)?;
    }
    unsafe {
        count.write(count_value);
        output.write(strings.finish()?.into_raw_parts().0);
    }
    Ok(())
}

fn push_pwstr(
    output: &mut CoTaskMemArrayBuilder<PWSTR, FreePwstrElements>,
    value: String,
) -> Result<()> {
    let value = OwnedPwstr::new(value)?;
    let raw = value.into_raw();
    if let Err(raw) = output.push(raw) {
        drop(unsafe { OwnedPwstr::from_raw(raw.0) });
        return Err(Error::from_hresult(E_UNEXPECTED));
    }
    Ok(())
}

fn decode_strings(count: u32, values: *const PCWSTR) -> Result<Vec<String>> {
    unsafe { borrow_input(values, count)? }
        .iter()
        .map(|value| {
            unsafe { value.to_string() }
                .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))
        })
        .collect()
}

fn checked_count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::from_hresult(E_UNEXPECTED))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::OPCAE_STATUS_RUNNING;
    use crate::client::AeClient;

    struct TestService;

    impl AeService for TestService {
        fn status(&self) -> Result<AeStatus> {
            Ok(AeStatus {
                start_time: FILETIME::default(),
                current_time: FILETIME::default(),
                last_update_time: FILETIME::default(),
                state: OPCAE_STATUS_RUNNING,
                version: (1, 2, 3),
                vendor_info: "rust-opc test".into(),
            })
        }

        fn available_filters(&self) -> Result<u32> {
            Ok(7)
        }

        fn event_categories(&self, _event_type: u32) -> Result<Vec<(u32, String)>> {
            Ok(vec![(10, "process".into())])
        }

        fn condition_names(&self, _event_category: u32) -> Result<Vec<String>> {
            Ok(vec!["high".into()])
        }

        fn subcondition_names(&self, _condition: &str) -> Result<Vec<String>> {
            Ok(vec!["high-high".into()])
        }

        fn source_conditions(&self, _source: &str) -> Result<Vec<String>> {
            Ok(vec!["high".into()])
        }

        fn event_attributes(&self, _event_category: u32) -> Result<Vec<(u32, String, u16)>> {
            Ok(vec![(20, "limit".into(), 5)])
        }

        fn enable_area(&self, _names: &[String], _enabled: bool) -> Result<()> {
            Ok(())
        }

        fn enable_source(&self, _names: &[String], _enabled: bool) -> Result<()> {
            Ok(())
        }
    }

    #[test]
    fn client_and_server_adapters_round_trip_task_memory() {
        let apartment = opc_classic_utils::ComApartment::mta().unwrap();
        let raw: IOPCEventServer = AeServerAdapter::new(Arc::new(TestService)).into();
        let client = AeClient::from_interface(&apartment, raw);

        assert_eq!(client.status().unwrap().vendor_info, "rust-opc test");
        assert_eq!(client.available_filters().unwrap(), 7);
        assert_eq!(
            client.event_categories(1).unwrap()[0].description,
            "process"
        );
        assert_eq!(client.condition_names(10).unwrap(), ["high"]);
        assert_eq!(client.subcondition_names("high").unwrap(), ["high-high"]);
        assert_eq!(client.event_attributes(10).unwrap()[0].description, "limit");
        client.enable_areas(&["plant"]).unwrap();
    }
}
