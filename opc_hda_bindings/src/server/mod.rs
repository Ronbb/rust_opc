//! Safe Rust service contracts and COM adapters for OPC Historical Data Access.

use std::sync::Arc;

use opc_classic_utils::server::{borrow_input, catch_ffi, initialize_output};
use opc_classic_utils::{
    Cleanup, CoTaskMemArrayBuilder, DropElements, FreePwstrElements, NoCleanup, OwnedPwstr,
};
use windows::Win32::Foundation::{
    E_INVALIDARG, E_NOTIMPL, E_POINTER, E_UNEXPECTED, FILETIME, S_FALSE,
};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Variant::VARIANT;
use windows_core::{Error, HRESULT, OutRef, PCWSTR, PWSTR, Result};

use crate::client::{
    HdaAggregate, HdaAttribute, HdaHistorianStatus, HdaServerHandle, HdaTime, HistoricalItemValues,
};
use crate::{
    IOPCHDA_Browser, IOPCHDA_Server, IOPCHDA_Server_Impl, IOPCHDA_SyncRead, IOPCHDA_SyncRead_Impl,
    tagOPCHDA_ATTRIBUTE, tagOPCHDA_ITEM, tagOPCHDA_MODIFIEDITEM, tagOPCHDA_OPERATORCODES,
    tagOPCHDA_TIME,
};

#[derive(Clone, Debug)]
pub struct HdaServerItemResult<T> {
    pub result: std::result::Result<T, HRESULT>,
}

impl<T> HdaServerItemResult<T> {
    pub fn success(value: T) -> Self {
        Self { result: Ok(value) }
    }

    pub fn failure(code: HRESULT) -> Self {
        Self { result: Err(code) }
    }
}

#[derive(Clone, Debug)]
pub struct HdaRawRead {
    pub resolved_start: FILETIME,
    pub resolved_end: FILETIME,
    pub items: Vec<HdaServerItemResult<HistoricalItemValues>>,
}

pub trait HdaService: Send + Sync + 'static {
    fn attributes(&self) -> Result<Vec<HdaAttribute>>;
    fn aggregates(&self) -> Result<Vec<HdaAggregate>>;
    fn status(&self) -> Result<HdaHistorianStatus>;
    fn item_handles(
        &self,
        item_ids: &[String],
        client_handles: &[u32],
    ) -> Result<Vec<HdaServerItemResult<HdaServerHandle>>>;
    fn release_item_handles(
        &self,
        handles: &[HdaServerHandle],
    ) -> Result<Vec<HdaServerItemResult<()>>>;
    fn validate_item_ids(&self, item_ids: &[String]) -> Result<Vec<HdaServerItemResult<()>>>;
    fn read_raw(
        &self,
        start: HdaTime,
        end: HdaTime,
        maximum_values: u32,
        include_bounds: bool,
        handles: &[HdaServerHandle],
    ) -> Result<HdaRawRead>;
    fn write_values(
        &self,
        _values: &[(HdaServerHandle, FILETIME, VARIANT, u32)],
    ) -> Result<Vec<HdaServerItemResult<()>>>;
}

#[derive(Clone)]
pub struct HdaServiceHandle(pub Arc<dyn HdaService>);

impl HdaServiceHandle {
    pub fn new(service: Arc<dyn HdaService>) -> Self {
        Self(service)
    }
}

#[windows_core::implement(IOPCHDA_Server, IOPCHDA_SyncRead)]
pub struct HdaServerAdapter {
    service: Arc<dyn HdaService>,
}

impl HdaServerAdapter {
    pub fn new(service: Arc<dyn HdaService>) -> Self {
        Self { service }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCHDA_Server_Impl for HdaServerAdapter_Impl {
    fn GetItemAttributes(
        &self,
        count: *mut u32,
        ids: *mut *mut u32,
        names: *mut *mut PWSTR,
        descriptions: *mut *mut PWSTR,
        data_types: *mut *mut u16,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(ids)?;
                initialize_output(names)?;
                initialize_output(descriptions)?;
                initialize_output(data_types)?;
            }
            let values = self.service.attributes()?;
            let count_value = checked_count(values.len())?;
            let mut id_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            let mut name_values = CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            let mut description_values =
                CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            let mut type_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            for value in values {
                push(&mut id_values, value.id)?;
                push_pwstr(&mut name_values, value.name)?;
                push_pwstr(&mut description_values, value.description)?;
                push(&mut type_values, value.data_type)?;
            }
            unsafe {
                count.write(count_value);
                ids.write(id_values.finish()?.into_raw_parts().0);
                names.write(name_values.finish()?.into_raw_parts().0);
                descriptions.write(description_values.finish()?.into_raw_parts().0);
                data_types.write(type_values.finish()?.into_raw_parts().0);
            }
            Ok(())
        })
    }

    fn GetAggregates(
        &self,
        count: *mut u32,
        ids: *mut *mut u32,
        names: *mut *mut PWSTR,
        descriptions: *mut *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(count)?;
                initialize_output(ids)?;
                initialize_output(names)?;
                initialize_output(descriptions)?;
            }
            let values = self.service.aggregates()?;
            let count_value = checked_count(values.len())?;
            let mut id_values = CoTaskMemArrayBuilder::new(values.len(), NoCleanup)?;
            let mut name_values = CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            let mut description_values =
                CoTaskMemArrayBuilder::new(values.len(), FreePwstrElements)?;
            for value in values {
                push(&mut id_values, value.id)?;
                push_pwstr(&mut name_values, value.name)?;
                push_pwstr(&mut description_values, value.description)?;
            }
            unsafe {
                count.write(count_value);
                ids.write(id_values.finish()?.into_raw_parts().0);
                names.write(name_values.finish()?.into_raw_parts().0);
                descriptions.write(description_values.finish()?.into_raw_parts().0);
            }
            Ok(())
        })
    }

    fn GetHistorianStatus(
        &self,
        state: *mut crate::tagOPCHDA_SERVERSTATUS,
        current_time: *mut *mut FILETIME,
        start_time: *mut *mut FILETIME,
        major: *mut u16,
        minor: *mut u16,
        build: *mut u16,
        max_return_values: *mut u32,
        status_text: *mut PWSTR,
        vendor_info: *mut PWSTR,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(state)?;
                initialize_output(current_time)?;
                initialize_output(start_time)?;
                initialize_output(major)?;
                initialize_output(minor)?;
                initialize_output(build)?;
                initialize_output(max_return_values)?;
                initialize_output(status_text)?;
                initialize_output(vendor_info)?;
            }
            let status = self.service.status()?;
            let mut current = CoTaskMemArrayBuilder::new(1, NoCleanup)?;
            push(&mut current, status.current_time)?;
            let mut start = CoTaskMemArrayBuilder::new(1, NoCleanup)?;
            push(&mut start, status.start_time)?;
            let status_string = OwnedPwstr::new(status.status)?;
            let vendor_string = OwnedPwstr::new(status.vendor_info)?;
            unsafe {
                state.write(status.state);
                current_time.write(current.finish()?.into_raw_parts().0);
                start_time.write(start.finish()?.into_raw_parts().0);
                major.write(status.version.0);
                minor.write(status.version.1);
                build.write(status.version.2);
                max_return_values.write(status.max_return_values);
                status_text.write(status_string.into_raw());
                vendor_info.write(vendor_string.into_raw());
            }
            Ok(())
        })
    }

    fn GetItemHandles(
        &self,
        count: u32,
        item_ids: *const PCWSTR,
        client_handles: *const u32,
        server_handles: *mut *mut u32,
        errors: *mut *mut HRESULT,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(server_handles)?;
                initialize_output(errors)?;
            }
            let item_ids = decode_strings(count, item_ids)?;
            let client_handles = unsafe { borrow_input(client_handles, count)? };
            let values = self.service.item_handles(&item_ids, client_handles)?;
            ensure_batch_len(values.len(), count as usize)?;
            let mut handles = CoTaskMemArrayBuilder::new(count as usize, NoCleanup)?;
            let mut item_errors = CoTaskMemArrayBuilder::new(count as usize, NoCleanup)?;
            let mut has_error = false;
            for value in values {
                match value.result {
                    Ok(handle) => {
                        push(&mut handles, handle.0)?;
                        push(&mut item_errors, HRESULT(0))?;
                    }
                    Err(error) => {
                        has_error = true;
                        push(&mut handles, 0)?;
                        push(&mut item_errors, error)?;
                    }
                }
            }
            unsafe {
                server_handles.write(handles.finish()?.into_raw_parts().0);
                errors.write(item_errors.finish()?.into_raw_parts().0);
            }
            batch_status(has_error)
        })
    }

    fn ReleaseItemHandles(
        &self,
        count: u32,
        server_handles: *const u32,
        errors: *mut *mut HRESULT,
    ) -> Result<()> {
        catch_ffi(|| {
            let handles = unsafe { borrow_input(server_handles, count)? }
                .iter()
                .copied()
                .map(HdaServerHandle)
                .collect::<Vec<_>>();
            write_unit_results(self.service.release_item_handles(&handles)?, count, errors)
        })
    }

    fn ValidateItemIDs(
        &self,
        count: u32,
        item_ids: *const PCWSTR,
        errors: *mut *mut HRESULT,
    ) -> Result<()> {
        catch_ffi(|| {
            let item_ids = decode_strings(count, item_ids)?;
            write_unit_results(self.service.validate_item_ids(&item_ids)?, count, errors)
        })
    }

    fn CreateBrowse(
        &self,
        _count: u32,
        _attribute_ids: *const u32,
        _operators: *const tagOPCHDA_OPERATORCODES,
        _filters: *const VARIANT,
        _browser: OutRef<IOPCHDA_Browser>,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IOPCHDA_SyncRead_Impl for HdaServerAdapter_Impl {
    fn ReadRaw(
        &self,
        start: *mut tagOPCHDA_TIME,
        end: *mut tagOPCHDA_TIME,
        maximum_values: u32,
        include_bounds: windows_core::BOOL,
        count: u32,
        server_handles: *const u32,
        item_values: *mut *mut tagOPCHDA_ITEM,
        errors: *mut *mut HRESULT,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe {
                initialize_output(item_values)?;
                initialize_output(errors)?;
            }
            let requested_start = decode_time(start)?;
            let requested_end = decode_time(end)?;
            let handles = unsafe { borrow_input(server_handles, count)? }
                .iter()
                .copied()
                .map(HdaServerHandle)
                .collect::<Vec<_>>();
            let result = self.service.read_raw(
                requested_start,
                requested_end,
                maximum_values,
                include_bounds.as_bool(),
                &handles,
            )?;
            ensure_batch_len(result.items.len(), count as usize)?;
            let mut values = CoTaskMemArrayBuilder::new(count as usize, HdaItemCleanup)?;
            let mut item_errors = CoTaskMemArrayBuilder::new(count as usize, NoCleanup)?;
            let mut has_error = false;
            for item in result.items {
                match item.result {
                    Ok(item) => {
                        push_hda_item(&mut values, build_hda_item(item)?)?;
                        push(&mut item_errors, HRESULT(0))?;
                    }
                    Err(error) => {
                        has_error = true;
                        push_hda_item(&mut values, tagOPCHDA_ITEM::default())?;
                        push(&mut item_errors, error)?;
                    }
                }
            }
            let values = values.finish()?;
            let item_errors = item_errors.finish()?;
            unsafe {
                (*start).bString = false.into();
                (*start).szTime = PWSTR::null();
                (*start).ftTime = result.resolved_start;
                (*end).bString = false.into();
                (*end).szTime = PWSTR::null();
                (*end).ftTime = result.resolved_end;
                item_values.write(values.into_raw_parts().0);
                errors.write(item_errors.into_raw_parts().0);
            }
            batch_status(has_error)
        })
    }

    fn ReadProcessed(
        &self,
        _start: *mut tagOPCHDA_TIME,
        _end: *mut tagOPCHDA_TIME,
        _resample_interval: &FILETIME,
        _count: u32,
        _server_handles: *const u32,
        _aggregates: *const u32,
        _item_values: *mut *mut tagOPCHDA_ITEM,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn ReadAtTime(
        &self,
        _timestamp_count: u32,
        _timestamps: *const FILETIME,
        _item_count: u32,
        _server_handles: *const u32,
        _item_values: *mut *mut tagOPCHDA_ITEM,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn ReadModified(
        &self,
        _start: *mut tagOPCHDA_TIME,
        _end: *mut tagOPCHDA_TIME,
        _maximum_values: u32,
        _count: u32,
        _server_handles: *const u32,
        _item_values: *mut *mut tagOPCHDA_MODIFIEDITEM,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }

    fn ReadAttribute(
        &self,
        _start: *mut tagOPCHDA_TIME,
        _end: *mut tagOPCHDA_TIME,
        _server_handle: u32,
        _attribute_count: u32,
        _attribute_ids: *const u32,
        _attribute_values: *mut *mut tagOPCHDA_ATTRIBUTE,
        _errors: *mut *mut HRESULT,
    ) -> Result<()> {
        Err(Error::from_hresult(E_NOTIMPL))
    }
}

struct HdaItemCleanup;

// SAFETY: Every initialized HDA item owns three independent task allocations.
unsafe impl Cleanup<tagOPCHDA_ITEM> for HdaItemCleanup {
    unsafe fn cleanup(&mut self, ptr: *mut tagOPCHDA_ITEM, initialized: usize) {
        for index in 0..initialized {
            let item = unsafe { &mut *ptr.add(index) };
            if !item.pvDataValues.is_null() {
                for value_index in 0..item.dwCount as usize {
                    unsafe { std::ptr::drop_in_place(item.pvDataValues.add(value_index)) };
                }
                unsafe { CoTaskMemFree(Some(item.pvDataValues.cast())) };
            }
            if !item.pftTimeStamps.is_null() {
                unsafe { CoTaskMemFree(Some(item.pftTimeStamps.cast())) };
            }
            if !item.pdwQualities.is_null() {
                unsafe { CoTaskMemFree(Some(item.pdwQualities.cast())) };
            }
            *item = tagOPCHDA_ITEM::default();
        }
    }
}

fn build_hda_item(item: HistoricalItemValues) -> Result<tagOPCHDA_ITEM> {
    let count = checked_count(item.samples.len())?;
    let mut timestamps = CoTaskMemArrayBuilder::new(item.samples.len(), NoCleanup)?;
    let mut qualities = CoTaskMemArrayBuilder::new(item.samples.len(), NoCleanup)?;
    let mut values = CoTaskMemArrayBuilder::new(item.samples.len(), DropElements)?;
    for sample in item.samples {
        push(&mut timestamps, sample.timestamp)?;
        push(&mut qualities, sample.quality)?;
        push(&mut values, sample.value)?;
    }
    let timestamps = timestamps.finish()?;
    let qualities = qualities.finish()?;
    let values = values.finish()?;
    Ok(tagOPCHDA_ITEM {
        hClient: item.client_handle.0,
        haAggregate: item.aggregate,
        dwCount: count,
        pftTimeStamps: timestamps.into_raw_parts().0,
        pdwQualities: qualities.into_raw_parts().0,
        pvDataValues: values.into_raw_parts().0,
    })
}

fn push_hda_item(
    output: &mut CoTaskMemArrayBuilder<tagOPCHDA_ITEM, HdaItemCleanup>,
    mut item: tagOPCHDA_ITEM,
) -> Result<()> {
    if let Err(returned) = output.push(item) {
        item = returned;
        unsafe { HdaItemCleanup.cleanup(&mut item, 1) };
        return Err(Error::from_hresult(E_UNEXPECTED));
    }
    Ok(())
}

fn write_unit_results(
    values: Vec<HdaServerItemResult<()>>,
    count: u32,
    errors: *mut *mut HRESULT,
) -> Result<()> {
    unsafe { initialize_output(errors)? };
    ensure_batch_len(values.len(), count as usize)?;
    let mut item_errors = CoTaskMemArrayBuilder::new(count as usize, NoCleanup)?;
    let mut has_error = false;
    for value in values {
        let error = match value.result {
            Ok(()) => HRESULT(0),
            Err(error) => {
                has_error = true;
                error
            }
        };
        push(&mut item_errors, error)?;
    }
    unsafe { errors.write(item_errors.finish()?.into_raw_parts().0) };
    batch_status(has_error)
}

fn decode_strings(count: u32, values: *const PCWSTR) -> Result<Vec<String>> {
    unsafe { borrow_input(values, count)? }
        .iter()
        .map(|value| unsafe { value.to_string() }.map_err(|_| Error::from_hresult(E_INVALIDARG)))
        .collect()
}

fn decode_time(value: *const tagOPCHDA_TIME) -> Result<HdaTime> {
    let value = unsafe { value.as_ref() }.ok_or_else(|| Error::from_hresult(E_POINTER))?;
    if value.bString.as_bool() {
        let expression =
            unsafe { value.szTime.to_string() }.map_err(|_| Error::from_hresult(E_INVALIDARG))?;
        Ok(HdaTime::Expression(expression))
    } else {
        Ok(HdaTime::Absolute(value.ftTime))
    }
}

fn push<T, C: Cleanup<T>>(output: &mut CoTaskMemArrayBuilder<T, C>, value: T) -> Result<()> {
    output
        .push(value)
        .map_err(|_| Error::from_hresult(E_UNEXPECTED))
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

fn ensure_batch_len(actual: usize, expected: usize) -> Result<()> {
    if actual == expected {
        Ok(())
    } else {
        Err(Error::from_hresult(E_UNEXPECTED))
    }
}

fn checked_count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::from_hresult(E_INVALIDARG))
}

fn batch_status(has_error: bool) -> Result<()> {
    if has_error {
        Err(Error::from_hresult(S_FALSE))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::OPCHDA_UP;
    use crate::client::{HdaClient, HdaClientHandle, HistoricalSample};

    struct TestService;

    impl HdaService for TestService {
        fn attributes(&self) -> Result<Vec<HdaAttribute>> {
            Ok(vec![HdaAttribute {
                id: 1,
                name: "Data type".into(),
                description: "Canonical data type".into(),
                data_type: 2,
            }])
        }

        fn aggregates(&self) -> Result<Vec<HdaAggregate>> {
            Ok(vec![HdaAggregate {
                id: 1,
                name: "Average".into(),
                description: "Time average".into(),
            }])
        }

        fn status(&self) -> Result<HdaHistorianStatus> {
            Ok(HdaHistorianStatus {
                state: OPCHDA_UP,
                current_time: FILETIME::default(),
                start_time: FILETIME::default(),
                version: (1, 2, 3),
                max_return_values: 100,
                status: "running".into(),
                vendor_info: "rust-opc test".into(),
            })
        }

        fn item_handles(
            &self,
            item_ids: &[String],
            _client_handles: &[u32],
        ) -> Result<Vec<HdaServerItemResult<HdaServerHandle>>> {
            Ok(item_ids
                .iter()
                .enumerate()
                .map(|(index, _)| HdaServerItemResult::success(HdaServerHandle(index as u32 + 10)))
                .collect())
        }

        fn release_item_handles(
            &self,
            handles: &[HdaServerHandle],
        ) -> Result<Vec<HdaServerItemResult<()>>> {
            Ok(handles
                .iter()
                .map(|_| HdaServerItemResult::success(()))
                .collect())
        }

        fn validate_item_ids(&self, item_ids: &[String]) -> Result<Vec<HdaServerItemResult<()>>> {
            Ok(item_ids
                .iter()
                .map(|_| HdaServerItemResult::success(()))
                .collect())
        }

        fn read_raw(
            &self,
            start: HdaTime,
            end: HdaTime,
            _maximum_values: u32,
            _include_bounds: bool,
            handles: &[HdaServerHandle],
        ) -> Result<HdaRawRead> {
            let HdaTime::Absolute(resolved_start) = start else {
                return Err(Error::from_hresult(E_INVALIDARG));
            };
            let HdaTime::Absolute(resolved_end) = end else {
                return Err(Error::from_hresult(E_INVALIDARG));
            };
            Ok(HdaRawRead {
                resolved_start,
                resolved_end,
                items: handles
                    .iter()
                    .map(|_| {
                        HdaServerItemResult::success(HistoricalItemValues {
                            client_handle: HdaClientHandle(7),
                            aggregate: 0,
                            samples: vec![HistoricalSample {
                                timestamp: FILETIME::default(),
                                quality: 192,
                                value: VARIANT::from(42i32),
                            }],
                        })
                    })
                    .collect(),
            })
        }

        fn write_values(
            &self,
            values: &[(HdaServerHandle, FILETIME, VARIANT, u32)],
        ) -> Result<Vec<HdaServerItemResult<()>>> {
            Ok(values
                .iter()
                .map(|_| HdaServerItemResult::success(()))
                .collect())
        }
    }

    #[test]
    fn client_and_server_adapters_round_trip_nested_task_memory() {
        let apartment = opc_classic_utils::ComApartment::mta().unwrap();
        let raw: IOPCHDA_Server = HdaServerAdapter::new(Arc::new(TestService)).into();
        let client = HdaClient::from_interface(&apartment, raw);

        assert_eq!(client.attributes().unwrap()[0].name, "Data type");
        assert_eq!(client.aggregates().unwrap()[0].name, "Average");
        assert_eq!(
            client.historian_status().unwrap().vendor_info,
            "rust-opc test"
        );

        let handles = client
            .item_handles(&["item"], &[HdaClientHandle(7)])
            .unwrap();
        let handle = handles[0].as_ref().unwrap().server_handle;
        let read = client
            .read_raw(
                HdaTime::Absolute(FILETIME::default()),
                HdaTime::Absolute(FILETIME::default()),
                10,
                false,
                &[handle],
            )
            .unwrap();
        let item = read.items[0].as_ref().unwrap();
        assert_eq!(item.samples.len(), 1);
        assert_eq!(item.samples[0].quality, 192);
        client.release_item_handles(&[handle]).unwrap();
    }
}
