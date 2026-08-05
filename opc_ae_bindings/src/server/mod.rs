//! Windows-independent Rust service contracts and a private COM adapter for
//! OPC Alarms & Events servers.

use std::sync::Arc;

use opc_classic_types::{ComObject, Error, ErrorCode, Result, Timestamp};
use opc_classic_utils::server::{borrow_input, catch_ffi, initialize_output};
use opc_classic_utils::{Cleanup, CoTaskMemArrayBuilder, FreePwstrElements, NoCleanup, OwnedPwstr};
use windows::Win32::Foundation::FILETIME;
use windows_core::{GUID, HRESULT, IUnknown, OutRef, PCWSTR, PWSTR, Result as AbiResult};

use crate::abi::{StatusCleanup, object_from_interface, to_abi_error, to_abi_timestamp};
use crate::{AeServerState, IOPCEventServer, IOPCEventServer_Impl};

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct AeStatus {
    pub start_time: Timestamp,
    pub current_time: Timestamp,
    pub last_update_time: Timestamp,
    pub state: AeServerState,
    pub version: (u16, u16, u16),
    pub vendor_info: String,
}

/// Application-facing AE server behavior. No method exposes a Windows type or
/// a Windows error.
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

#[derive(Clone)]
pub struct AeServiceHandle(pub Arc<dyn AeService>);

impl AeServiceHandle {
    pub fn new(service: Arc<dyn AeService>) -> Self {
        Self(service)
    }
}

/// Per-event status represented by the project-owned error code.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct EventError(pub ErrorCode);

/// An owning COM object backed by a Rust [`AeService`].
///
/// The generated ABI implementation stays private to this module; consumers
/// exchange only [`ComObject`] values.
pub struct AeServer {
    object: ComObject,
}

impl AeServer {
    pub fn new(service: Arc<dyn AeService>) -> Self {
        let interface: IOPCEventServer = AeServerAdapter { service }.into();
        Self {
            object: object_from_interface(&interface),
        }
    }

    pub fn object(&self) -> ComObject {
        self.object.clone()
    }
}

#[windows_core::implement(IOPCEventServer)]
struct AeServerAdapter {
    service: Arc<dyn AeService>,
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCEventServer_Impl for AeServerAdapter_Impl {
    fn GetStatus(&self) -> AbiResult<*mut crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005> {
        service_call(|| {
            let status = self.service.status()?;
            let vendor = OwnedPwstr::new(status.vendor_info)?;
            let mut output = CoTaskMemArrayBuilder::new(1, StatusCleanup)?;
            let value = crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0005 {
                ftStartTime: to_abi_timestamp(status.start_time),
                ftCurrentTime: to_abi_timestamp(status.current_time),
                ftLastUpdateTime: to_abi_timestamp(status.last_update_time),
                dwServerState: crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0003(status.state.raw()),
                wMajorVersion: status.version.0,
                wMinorVersion: status.version.1,
                wBuildNumber: status.version.2,
                wReserved: 0,
                szVendorInfo: PWSTR(vendor.into_raw()),
            };
            if let Err(mut value) = output.push(value) {
                unsafe { StatusCleanup.cleanup(&mut value, 1) };
                return Err(Error::unexpected("AE status output is already initialized"));
            }
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
    ) -> AbiResult<()> {
        not_implemented("AE event subscriptions are not implemented")
    }

    fn QueryAvailableFilters(&self) -> AbiResult<u32> {
        service_call(|| self.service.available_filters())
    }

    fn QueryEventCategories(
        &self,
        event_type: u32,
        count: *mut u32,
        ids: *mut *mut u32,
        descriptions: *mut *mut PWSTR,
    ) -> AbiResult<()> {
        service_call(|| {
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
                    .map_err(|_| Error::unexpected("AE category output exceeds its capacity"))?;
                push_pwstr(&mut description_values, description)?;
            }
            let id_values = id_values.finish()?;
            let description_values = description_values.finish()?;
            let (id_ptr, _) = id_values.into_raw_parts();
            let (description_ptr, _) = description_values.into_raw_parts();
            unsafe {
                count.write(count_value);
                ids.write(id_ptr);
                descriptions.write(description_ptr.cast::<PWSTR>());
            }
            Ok(())
        })
    }

    fn QueryConditionNames(
        &self,
        category: u32,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> AbiResult<()> {
        service_call(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(names)?;
            }
            write_strings(self.service.condition_names(category)?, count, names)
        })
    }

    fn QuerySubConditionNames(
        &self,
        condition: &PCWSTR,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> AbiResult<()> {
        service_call(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(names)?;
            }
            if condition.is_null() {
                return Err(Error::null_pointer("null condition name"));
            }
            let condition = unsafe { condition.to_string() }
                .map_err(|_| Error::invalid_argument("condition is not valid UTF-16"))?;
            write_strings(self.service.subcondition_names(&condition)?, count, names)
        })
    }

    fn QuerySourceConditions(
        &self,
        source: &PCWSTR,
        count: *mut u32,
        names: *mut *mut PWSTR,
    ) -> AbiResult<()> {
        service_call(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(names)?;
            }
            if source.is_null() {
                return Err(Error::null_pointer("null source name"));
            }
            let source = unsafe { source.to_string() }
                .map_err(|_| Error::invalid_argument("source is not valid UTF-16"))?;
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
    ) -> AbiResult<()> {
        service_call(|| {
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
                    .map_err(|_| Error::unexpected("AE attribute output exceeds its capacity"))?;
                push_pwstr(&mut description_values, description)?;
                type_values
                    .push(data_type)
                    .map_err(|_| Error::unexpected("AE type output exceeds its capacity"))?;
            }
            let id_values = id_values.finish()?;
            let description_values = description_values.finish()?;
            let type_values = type_values.finish()?;
            let (id_ptr, _) = id_values.into_raw_parts();
            let (description_ptr, _) = description_values.into_raw_parts();
            let (type_ptr, _) = type_values.into_raw_parts();
            unsafe {
                count.write(count_value);
                ids.write(id_ptr);
                descriptions.write(description_ptr.cast::<PWSTR>());
                data_types.write(type_ptr);
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
    ) -> AbiResult<()> {
        not_implemented("AE item translation is not implemented")
    }

    fn GetConditionState(
        &self,
        _source: &PCWSTR,
        _condition_name: &PCWSTR,
        _count: u32,
        _attribute_ids: *const u32,
    ) -> AbiResult<*mut crate::__MIDL___MIDL_itf_opc_ae_0000_0001_0006> {
        not_implemented("AE condition state is not implemented")
    }

    fn EnableConditionByArea(&self, count: u32, names: *const PCWSTR) -> AbiResult<()> {
        service_call(|| {
            self.service
                .enable_area(&decode_strings(count, names)?, true)
        })
    }

    fn EnableConditionBySource(&self, count: u32, names: *const PCWSTR) -> AbiResult<()> {
        service_call(|| {
            self.service
                .enable_source(&decode_strings(count, names)?, true)
        })
    }

    fn DisableConditionByArea(&self, count: u32, names: *const PCWSTR) -> AbiResult<()> {
        service_call(|| {
            self.service
                .enable_area(&decode_strings(count, names)?, false)
        })
    }

    fn DisableConditionBySource(&self, count: u32, names: *const PCWSTR) -> AbiResult<()> {
        service_call(|| {
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
    ) -> AbiResult<()> {
        not_implemented("AE condition acknowledgement is not implemented")
    }

    fn CreateAreaBrowser(&self, _iid: *const GUID) -> AbiResult<IUnknown> {
        not_implemented("AE area browsing is not implemented")
    }
}

fn service_call<T>(call: impl FnOnce() -> Result<T>) -> AbiResult<T> {
    catch_ffi(call).map_err(to_abi_error)
}

fn not_implemented<T>(message: &'static str) -> AbiResult<T> {
    Err(to_abi_error(Error::not_implemented(message)))
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
    let strings = strings.finish()?;
    let (strings_ptr, _) = strings.into_raw_parts();
    unsafe {
        count.write(count_value);
        output.write(strings_ptr.cast::<PWSTR>());
    }
    Ok(())
}

fn push_pwstr(
    output: &mut CoTaskMemArrayBuilder<*mut u16, FreePwstrElements>,
    value: String,
) -> Result<()> {
    let value = OwnedPwstr::new(value)?;
    let raw = value.into_raw();
    if let Err(raw) = output.push(raw) {
        drop(unsafe { OwnedPwstr::from_raw(raw) });
        return Err(Error::unexpected("AE string output exceeds its capacity"));
    }
    Ok(())
}

fn decode_strings(count: u32, values: *const PCWSTR) -> Result<Vec<String>> {
    unsafe { borrow_input(values, count)? }
        .iter()
        .map(|value| {
            if value.is_null() {
                return Err(Error::null_pointer("null AE string input"));
            }
            unsafe { value.to_string() }
                .map_err(|_| Error::invalid_argument("string is not valid UTF-16"))
        })
        .collect()
}

fn checked_count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::out_of_memory("AE output has too many elements"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::client::AeClient;

    struct TestService;

    impl AeService for TestService {
        fn status(&self) -> Result<AeStatus> {
            Ok(AeStatus {
                start_time: Timestamp::default(),
                current_time: Timestamp::default(),
                last_update_time: Timestamp::default(),
                state: AeServerState::Running,
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
    fn client_and_server_round_trip_without_public_windows_types() {
        let apartment = opc_classic_utils::ComApartment::mta().unwrap();
        let server = AeServer::new(Arc::new(TestService));
        let client = AeClient::from_object(&apartment, &server.object()).unwrap();

        let status = client.status().unwrap();
        assert_eq!(status.vendor_info, "rust-opc test");
        assert_eq!(status.state, AeServerState::Running);
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
