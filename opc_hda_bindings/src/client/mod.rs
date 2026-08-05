//! Safe OPC Historical Data Access client facade.

use std::marker::PhantomData;

use opc_classic_types::{
    ClassContext, ComObject, Error, ErrorCode, Guid, Result, Timestamp, Value, ValueType,
};
use opc_classic_utils::{
    Cleanup, CoTaskMemArrayOut, ComApartment, FreePwstrElements, NoCleanup, OwnedPwstr, WideCString,
};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::Foundation::FILETIME;
use windows::Win32::System::Com::{CLSCTX, CLSIDFromProgID, CoCreateInstance, CoTaskMemFree};
use windows_core::{HRESULT, IUnknown, PCWSTR};

use crate::convert::{
    from_abi_error, guid_to_abi, interface_from_object, object_from_interface, timestamp_from_abi,
    value_from_abi,
};
use crate::{
    IOPCHDA_Server, IOPCHDA_SyncRead, tagOPCHDA_ITEM, tagOPCHDA_SERVERSTATUS, tagOPCHDA_TIME,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct HdaServerHandle(pub u32);

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct HdaClientHandle(pub u32);

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct HdaItemError(pub ErrorCode);

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HdaAttribute {
    pub id: u32,
    pub name: String,
    pub description: String,
    pub data_type: ValueType,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HdaAggregate {
    pub id: u32,
    pub name: String,
    pub description: String,
}

#[derive(Clone, Debug)]
pub struct HdaHistorianStatus {
    pub state: HdaServerState,
    pub current_time: Timestamp,
    pub start_time: Timestamp,
    pub version: (u16, u16, u16),
    pub max_return_values: u32,
    pub status: String,
    pub vendor_info: String,
}

#[derive(Clone, Debug)]
pub struct HdaItem {
    pub client_handle: HdaClientHandle,
    pub server_handle: HdaServerHandle,
}

#[derive(Clone, Debug)]
pub struct HistoricalSample {
    pub timestamp: Timestamp,
    pub quality: u32,
    pub value: Value,
}

#[derive(Clone, Debug)]
pub struct HistoricalItemValues {
    pub client_handle: HdaClientHandle,
    pub aggregate: u32,
    pub samples: Vec<HistoricalSample>,
}

#[derive(Clone, Debug)]
pub enum HdaTime {
    Absolute(Timestamp),
    Expression(String),
}

#[derive(Clone, Debug)]
pub struct HdaReadResult {
    pub resolved_start: Timestamp,
    pub resolved_end: Timestamp,
    pub items: Vec<std::result::Result<HistoricalItemValues, HdaItemError>>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum HdaServerState {
    Up,
    Down,
    Indeterminate,
    Unknown(i32),
}

pub struct HdaClient<'apartment> {
    inner: IOPCHDA_Server,
    _apartment: PhantomData<&'apartment ComApartment>,
}

impl<'apartment> HdaClient<'apartment> {
    pub fn connect_prog_id(apartment: &'apartment ComApartment, prog_id: &str) -> Result<Self> {
        let prog_id = wide(prog_id)?;
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
        _apartment: &'apartment ComApartment,
        class_id: &Guid,
        context: ClassContext,
    ) -> Result<Self> {
        let class_id = guid_to_abi(class_id);
        let inner =
            unsafe { CoCreateInstance(&class_id, None::<&IUnknown>, CLSCTX(context.bits())) }
                .map_err(from_abi_error)?;
        Ok(Self {
            inner,
            _apartment: PhantomData,
        })
    }

    fn from_interface(apartment: &'apartment ComApartment, inner: IOPCHDA_Server) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn from_object(apartment: &'apartment ComApartment, object: &ComObject) -> Result<Self> {
        Ok(Self::from_interface(
            apartment,
            interface_from_object(object)?,
        ))
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(self.inner.clone())
    }

    pub fn common(&self) -> Result<CommonClient> {
        CommonClient::from_object(&self.object())
    }

    pub fn attributes(&self) -> Result<Vec<HdaAttribute>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut names = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut descriptions = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut data_types = CoTaskMemArrayOut::new(0, NoCleanup);
        let call = unsafe {
            self.inner.GetItemAttributes(
                &mut count,
                ids.as_mut_ptr(),
                names.as_mut_ptr().cast(),
                descriptions.as_mut_ptr().cast(),
                data_types.as_mut_ptr(),
            )
        };
        let len = count as usize;
        unsafe {
            ids.set_len(len);
            names.set_len(len);
            descriptions.set_len(len);
            data_types.set_len(len);
        }
        let ids = unsafe { ids.into_array() };
        let names = unsafe { names.into_array() };
        let descriptions = unsafe { descriptions.into_array() };
        let data_types = unsafe { data_types.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let names = names?;
        let descriptions = descriptions?;
        let data_types = data_types?;
        Ok((0..len)
            .map(|index| HdaAttribute {
                id: ids.as_slice()[index],
                name: pwstr_string(names.as_slice()[index]),
                description: pwstr_string(descriptions.as_slice()[index]),
                data_type: ValueType::from_raw(data_types.as_slice()[index]),
            })
            .collect())
    }

    pub fn aggregates(&self) -> Result<Vec<HdaAggregate>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut names = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut descriptions = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let call = unsafe {
            self.inner.GetAggregates(
                &mut count,
                ids.as_mut_ptr(),
                names.as_mut_ptr().cast(),
                descriptions.as_mut_ptr().cast(),
            )
        };
        let len = count as usize;
        unsafe {
            ids.set_len(len);
            names.set_len(len);
            descriptions.set_len(len);
        }
        let ids = unsafe { ids.into_array() };
        let names = unsafe { names.into_array() };
        let descriptions = unsafe { descriptions.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let names = names?;
        let descriptions = descriptions?;
        Ok((0..len)
            .map(|index| HdaAggregate {
                id: ids.as_slice()[index],
                name: pwstr_string(names.as_slice()[index]),
                description: pwstr_string(descriptions.as_slice()[index]),
            })
            .collect())
    }

    pub fn historian_status(&self) -> Result<HdaHistorianStatus> {
        let mut state = tagOPCHDA_SERVERSTATUS::default();
        let mut current = CoTaskMemArrayOut::new(1, NoCleanup);
        let mut start = CoTaskMemArrayOut::new(1, NoCleanup);
        let mut major = 0;
        let mut minor = 0;
        let mut build = 0;
        let mut max_return_values = 0;
        let mut status = windows_core::PWSTR::null();
        let mut vendor = windows_core::PWSTR::null();
        let call = unsafe {
            self.inner.GetHistorianStatus(
                &mut state,
                current.as_mut_ptr(),
                start.as_mut_ptr(),
                &mut major,
                &mut minor,
                &mut build,
                &mut max_return_values,
                &mut status,
                &mut vendor,
            )
        };
        let status = unsafe { OwnedPwstr::from_raw(status.0) };
        let vendor = unsafe { OwnedPwstr::from_raw(vendor.0) };
        let current = unsafe { current.into_array() };
        let start = unsafe { start.into_array() };
        call.map_err(from_abi_error)?;
        let current = current?;
        let start = start?;
        if current.is_empty() || start.is_empty() {
            return Err(Error::null_pointer(
                "historian status omitted required timestamps",
            ));
        }
        Ok(HdaHistorianStatus {
            state: state_from_abi(state),
            current_time: timestamp_from_abi(current.as_slice()[0]),
            start_time: timestamp_from_abi(start.as_slice()[0]),
            version: (major, minor, build),
            max_return_values,
            status: status.to_string_lossy(),
            vendor_info: vendor.to_string_lossy(),
        })
    }

    pub fn item_handles(
        &self,
        item_ids: &[&str],
        client_handles: &[HdaClientHandle],
    ) -> Result<Vec<std::result::Result<HdaItem, HdaItemError>>> {
        if item_ids.len() != client_handles.len() {
            return Err(Error::invalid_argument(
                "item IDs and client handles have different lengths",
            ));
        }
        let ids = item_ids
            .iter()
            .map(|id| wide(id))
            .collect::<Result<Vec<_>>>()?;
        let pointers: Vec<PCWSTR> = ids.iter().map(|id| PCWSTR(id.as_ptr())).collect();
        let clients: Vec<u32> = client_handles.iter().map(|handle| handle.0).collect();
        let mut servers = CoTaskMemArrayOut::new(ids.len(), NoCleanup);
        let mut errors = CoTaskMemArrayOut::new(ids.len(), NoCleanup);
        let call = unsafe {
            self.inner.GetItemHandles(
                count(ids.len())?,
                pointers.as_ptr(),
                clients.as_ptr(),
                servers.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let servers = unsafe { servers.into_array() };
        let errors = unsafe { errors.into_array() };
        call.map_err(from_abi_error)?;
        let servers = servers?;
        let errors = errors?;
        Ok((0..ids.len())
            .map(|index| {
                let error = errors.as_slice()[index];
                if error.is_err() {
                    Err(HdaItemError(ErrorCode::from_raw(error.0)))
                } else {
                    Ok(HdaItem {
                        client_handle: client_handles[index],
                        server_handle: HdaServerHandle(servers.as_slice()[index]),
                    })
                }
            })
            .collect())
    }

    pub fn release_item_handles(
        &self,
        handles: &[HdaServerHandle],
    ) -> Result<Vec<std::result::Result<(), HdaItemError>>> {
        let raw: Vec<u32> = handles.iter().map(|handle| handle.0).collect();
        let count = count(handles.len())?;
        self.item_errors(handles.len(), |errors| unsafe {
            self.inner.ReleaseItemHandles(count, raw.as_ptr(), errors)
        })
    }

    pub fn validate_item_ids(
        &self,
        item_ids: &[&str],
    ) -> Result<Vec<std::result::Result<(), HdaItemError>>> {
        let ids = item_ids
            .iter()
            .map(|id| wide(id))
            .collect::<Result<Vec<_>>>()?;
        let pointers: Vec<PCWSTR> = ids.iter().map(|id| PCWSTR(id.as_ptr())).collect();
        let count = count(ids.len())?;
        self.item_errors(ids.len(), |errors| unsafe {
            self.inner.ValidateItemIDs(count, pointers.as_ptr(), errors)
        })
    }

    pub fn read_raw(
        &self,
        start: HdaTime,
        end: HdaTime,
        maximum_values: u32,
        include_bounds: bool,
        handles: &[HdaServerHandle],
    ) -> Result<HdaReadResult> {
        let reader: IOPCHDA_SyncRead = interface_from_object(&self.object())?;
        let mut start = AbiTime::new(start)?;
        let mut end = AbiTime::new(end)?;
        let raw_handles: Vec<u32> = handles.iter().map(|handle| handle.0).collect();
        let mut values = CoTaskMemArrayOut::new(handles.len(), HdaItemCleanup);
        let mut errors = CoTaskMemArrayOut::new(handles.len(), NoCleanup);
        let call = unsafe {
            reader.ReadRaw(
                &mut start.raw,
                &mut end.raw,
                maximum_values,
                include_bounds,
                count(handles.len())?,
                raw_handles.as_ptr(),
                values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let values = unsafe { values.into_array() };
        let errors = unsafe { errors.into_array() };
        call.map_err(from_abi_error)?;
        let values = values?;
        let errors = errors?;
        let items = values
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(item, error)| -> Result<_> {
                if error.is_err() {
                    Ok(Err(HdaItemError(ErrorCode::from_raw(error.0))))
                } else {
                    decode_hda_item(item).map(Ok)
                }
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(HdaReadResult {
            resolved_start: timestamp_from_abi(start.raw.ftTime),
            resolved_end: timestamp_from_abi(end.raw.ftTime),
            items,
        })
    }

    fn item_errors(
        &self,
        len: usize,
        call: impl FnOnce(*mut *mut HRESULT) -> windows_core::Result<()>,
    ) -> Result<Vec<std::result::Result<(), HdaItemError>>> {
        let mut errors = CoTaskMemArrayOut::new(len, NoCleanup);
        let result = call(errors.as_mut_ptr());
        let errors = unsafe { errors.into_array() };
        result.map_err(from_abi_error)?;
        let errors = errors?;
        Ok(errors
            .as_slice()
            .iter()
            .map(|error| {
                if error.is_err() {
                    Err(HdaItemError(ErrorCode::from_raw(error.0)))
                } else {
                    Ok(())
                }
            })
            .collect())
    }
}

struct AbiTime {
    raw: tagOPCHDA_TIME,
    _expression: Option<WideCString>,
}

impl AbiTime {
    fn new(value: HdaTime) -> Result<Self> {
        match value {
            HdaTime::Absolute(time) => Ok(Self {
                raw: tagOPCHDA_TIME {
                    bString: false.into(),
                    szTime: windows_core::PWSTR::null(),
                    ftTime: crate::convert::timestamp_to_abi(time),
                },
                _expression: None,
            }),
            HdaTime::Expression(expression) => {
                let expression = wide(&expression)?;
                let raw = tagOPCHDA_TIME {
                    bString: true.into(),
                    szTime: windows_core::PWSTR(expression.as_ptr().cast_mut()),
                    ftTime: FILETIME::default(),
                };
                Ok(Self {
                    raw,
                    _expression: Some(expression),
                })
            }
        }
    }
}

struct HdaItemCleanup;

// SAFETY: Each OPCHDA_ITEM owns three task-allocated arrays; every variant in
// the value array owns its own Automation resources.
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

fn decode_hda_item(item: &tagOPCHDA_ITEM) -> Result<HistoricalItemValues> {
    let len = item.dwCount as usize;
    if len == 0 {
        return Ok(HistoricalItemValues {
            client_handle: HdaClientHandle(item.hClient),
            aggregate: item.haAggregate,
            samples: Vec::new(),
        });
    }
    if len != 0
        && (item.pftTimeStamps.is_null()
            || item.pdwQualities.is_null()
            || item.pvDataValues.is_null())
    {
        return Err(Error::null_pointer(
            "historical item omitted one or more sample arrays",
        ));
    }
    let timestamps = unsafe { std::slice::from_raw_parts(item.pftTimeStamps, len) };
    let qualities = unsafe { std::slice::from_raw_parts(item.pdwQualities, len) };
    let values = unsafe { std::slice::from_raw_parts(item.pvDataValues, len) };
    Ok(HistoricalItemValues {
        client_handle: HdaClientHandle(item.hClient),
        aggregate: item.haAggregate,
        samples: (0..len)
            .map(|index| {
                Ok(HistoricalSample {
                    timestamp: timestamp_from_abi(timestamps[index]),
                    quality: qualities[index],
                    value: value_from_abi(&values[index])?,
                })
            })
            .collect::<Result<Vec<_>>>()?,
    })
}

fn wide(value: &str) -> Result<WideCString> {
    WideCString::try_from(value).map_err(|_| Error::invalid_argument("string contains NUL"))
}

fn pwstr_string(value: *mut u16) -> String {
    if value.is_null() {
        String::new()
    } else {
        unsafe { windows_core::PWSTR(value).to_string() }.unwrap_or_default()
    }
}

fn count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::invalid_argument("item count exceeds u32"))
}

fn state_from_abi(value: tagOPCHDA_SERVERSTATUS) -> HdaServerState {
    match value {
        crate::OPCHDA_UP => HdaServerState::Up,
        crate::OPCHDA_DOWN => HdaServerState::Down,
        crate::OPCHDA_INDETERMINATE => HdaServerState::Indeterminate,
        other => HdaServerState::Unknown(other.0),
    }
}
