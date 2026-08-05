//! Safe OPC Historical Data Access client facade.

use std::marker::PhantomData;

use opc_classic_utils::{
    Cleanup, CoTaskMemOut, ComApartment, FreePwstrElements, NoCleanup, OwnedPwstr, WideCString,
};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::Foundation::{E_INVALIDARG, E_POINTER, FILETIME};
use windows::Win32::System::Com::{
    CLSCTX, CLSCTX_ALL, CLSIDFromProgID, CoCreateInstance, CoTaskMemFree,
};
use windows::Win32::System::Variant::VARIANT;
use windows_core::{Error, GUID, HRESULT, IUnknown, Interface, PCWSTR, Result};

use crate::{
    IOPCHDA_Server, IOPCHDA_SyncRead, tagOPCHDA_ITEM, tagOPCHDA_SERVERSTATUS, tagOPCHDA_TIME,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct HdaServerHandle(pub u32);

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct HdaClientHandle(pub u32);

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct HdaItemError(pub HRESULT);

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HdaAttribute {
    pub id: u32,
    pub name: String,
    pub description: String,
    pub data_type: u16,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HdaAggregate {
    pub id: u32,
    pub name: String,
    pub description: String,
}

#[derive(Clone, Debug)]
pub struct HdaHistorianStatus {
    pub state: tagOPCHDA_SERVERSTATUS,
    pub current_time: FILETIME,
    pub start_time: FILETIME,
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
    pub timestamp: FILETIME,
    pub quality: u32,
    pub value: VARIANT,
}

#[derive(Clone, Debug)]
pub struct HistoricalItemValues {
    pub client_handle: HdaClientHandle,
    pub aggregate: u32,
    pub samples: Vec<HistoricalSample>,
}

#[derive(Clone, Debug)]
pub enum HdaTime {
    Absolute(FILETIME),
    Expression(String),
}

#[derive(Clone, Debug)]
pub struct HdaReadResult {
    pub resolved_start: FILETIME,
    pub resolved_end: FILETIME,
    pub items: Vec<std::result::Result<HistoricalItemValues, HdaItemError>>,
}

pub struct HdaClient<'apartment> {
    inner: IOPCHDA_Server,
    _apartment: PhantomData<&'apartment ComApartment>,
}

impl<'apartment> HdaClient<'apartment> {
    pub fn connect_prog_id(apartment: &'apartment ComApartment, prog_id: &str) -> Result<Self> {
        let prog_id = wide(prog_id)?;
        let class_id = unsafe { CLSIDFromProgID(prog_id.as_pcwstr()) }?;
        Self::connect_class_id(apartment, &class_id, CLSCTX_ALL)
    }

    pub fn connect_class_id(
        _apartment: &'apartment ComApartment,
        class_id: &GUID,
        context: CLSCTX,
    ) -> Result<Self> {
        let inner = unsafe { CoCreateInstance(class_id, None::<&IUnknown>, context) }?;
        Ok(Self {
            inner,
            _apartment: PhantomData,
        })
    }

    pub fn from_interface(apartment: &'apartment ComApartment, inner: IOPCHDA_Server) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn common(&self) -> Result<CommonClient> {
        self.inner.cast().map(CommonClient::new)
    }

    pub fn attributes(&self) -> Result<Vec<HdaAttribute>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemOut::<u32>::new();
        let mut names = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut descriptions = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut data_types = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            self.inner.GetItemAttributes(
                &mut count,
                ids.as_mut_ptr(),
                names.as_mut_ptr(),
                descriptions.as_mut_ptr(),
                data_types.as_mut_ptr(),
            )
        };
        let len = count as usize;
        let ids = unsafe { ids.into_array(len, NoCleanup) }?;
        let names = unsafe { names.into_array(len, FreePwstrElements) }?;
        let descriptions = unsafe { descriptions.into_array(len, FreePwstrElements) }?;
        let data_types = unsafe { data_types.into_array(len, NoCleanup) }?;
        call?;
        Ok((0..len)
            .map(|index| HdaAttribute {
                id: ids.as_slice()[index],
                name: pwstr_string(names.as_slice()[index]),
                description: pwstr_string(descriptions.as_slice()[index]),
                data_type: data_types.as_slice()[index],
            })
            .collect())
    }

    pub fn aggregates(&self) -> Result<Vec<HdaAggregate>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemOut::<u32>::new();
        let mut names = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut descriptions = CoTaskMemOut::<windows_core::PWSTR>::new();
        let call = unsafe {
            self.inner.GetAggregates(
                &mut count,
                ids.as_mut_ptr(),
                names.as_mut_ptr(),
                descriptions.as_mut_ptr(),
            )
        };
        let len = count as usize;
        let ids = unsafe { ids.into_array(len, NoCleanup) }?;
        let names = unsafe { names.into_array(len, FreePwstrElements) }?;
        let descriptions = unsafe { descriptions.into_array(len, FreePwstrElements) }?;
        call?;
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
        let mut current = CoTaskMemOut::<FILETIME>::new();
        let mut start = CoTaskMemOut::<FILETIME>::new();
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
        let current_len = if current.is_null() { 0 } else { 1 };
        let start_len = if start.is_null() { 0 } else { 1 };
        let current = unsafe { current.into_array(current_len, NoCleanup) }?;
        let start = unsafe { start.into_array(start_len, NoCleanup) }?;
        let status = unsafe { OwnedPwstr::from_raw(status.0) };
        let vendor = unsafe { OwnedPwstr::from_raw(vendor.0) };
        call?;
        if current.is_empty() || start.is_empty() {
            return Err(Error::from_hresult(E_POINTER));
        }
        Ok(HdaHistorianStatus {
            state,
            current_time: current.as_slice()[0],
            start_time: start.as_slice()[0],
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
            return Err(Error::from_hresult(E_INVALIDARG));
        }
        let ids = item_ids
            .iter()
            .map(|id| wide(id))
            .collect::<Result<Vec<_>>>()?;
        let pointers: Vec<PCWSTR> = ids.iter().map(WideCString::as_pcwstr).collect();
        let clients: Vec<u32> = client_handles.iter().map(|handle| handle.0).collect();
        let mut servers = CoTaskMemOut::<u32>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            self.inner.GetItemHandles(
                count(ids.len())?,
                pointers.as_ptr(),
                clients.as_ptr(),
                servers.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let servers = unsafe { servers.into_array(ids.len(), NoCleanup) }?;
        let errors = unsafe { errors.into_array(ids.len(), NoCleanup) }?;
        call?;
        Ok((0..ids.len())
            .map(|index| {
                let error = errors.as_slice()[index];
                if error.is_err() {
                    Err(HdaItemError(error))
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
        self.item_errors(handles.len(), |errors| unsafe {
            self.inner
                .ReleaseItemHandles(count(handles.len())?, raw.as_ptr(), errors)
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
        let pointers: Vec<PCWSTR> = ids.iter().map(WideCString::as_pcwstr).collect();
        self.item_errors(ids.len(), |errors| unsafe {
            self.inner
                .ValidateItemIDs(count(ids.len())?, pointers.as_ptr(), errors)
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
        let reader: IOPCHDA_SyncRead = self.inner.cast()?;
        let mut start = AbiTime::new(start)?;
        let mut end = AbiTime::new(end)?;
        let raw_handles: Vec<u32> = handles.iter().map(|handle| handle.0).collect();
        let mut values = CoTaskMemOut::<tagOPCHDA_ITEM>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
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
        let values = unsafe { values.into_array(handles.len(), HdaItemCleanup) }?;
        let errors = unsafe { errors.into_array(handles.len(), NoCleanup) }?;
        call?;
        let items = values
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(item, error)| -> Result<_> {
                if error.is_err() {
                    Ok(Err(HdaItemError(*error)))
                } else {
                    decode_hda_item(item).map(Ok)
                }
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(HdaReadResult {
            resolved_start: start.raw.ftTime,
            resolved_end: end.raw.ftTime,
            items,
        })
    }

    pub fn as_raw(&self) -> &IOPCHDA_Server {
        &self.inner
    }

    fn item_errors(
        &self,
        len: usize,
        call: impl FnOnce(*mut *mut HRESULT) -> Result<()>,
    ) -> Result<Vec<std::result::Result<(), HdaItemError>>> {
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let result = call(errors.as_mut_ptr());
        let errors = unsafe { errors.into_array(len, NoCleanup) }?;
        result?;
        Ok(errors
            .as_slice()
            .iter()
            .map(|error| {
                if error.is_err() {
                    Err(HdaItemError(*error))
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
                    ftTime: time,
                },
                _expression: None,
            }),
            HdaTime::Expression(expression) => {
                let expression = wide(&expression)?;
                let raw = tagOPCHDA_TIME {
                    bString: true.into(),
                    szTime: windows_core::PWSTR(expression.as_pcwstr().0.cast_mut()),
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
        return Err(Error::from_hresult(E_POINTER));
    }
    let timestamps = unsafe { std::slice::from_raw_parts(item.pftTimeStamps, len) };
    let qualities = unsafe { std::slice::from_raw_parts(item.pdwQualities, len) };
    let values = unsafe { std::slice::from_raw_parts(item.pvDataValues, len) };
    Ok(HistoricalItemValues {
        client_handle: HdaClientHandle(item.hClient),
        aggregate: item.haAggregate,
        samples: (0..len)
            .map(|index| HistoricalSample {
                timestamp: timestamps[index],
                quality: qualities[index],
                value: values[index].clone(),
            })
            .collect(),
    })
}

fn wide(value: &str) -> Result<WideCString> {
    WideCString::try_from(value).map_err(|_| Error::from_hresult(E_INVALIDARG))
}

fn pwstr_string(value: windows_core::PWSTR) -> String {
    if value.is_null() {
        String::new()
    } else {
        unsafe { value.to_string() }.unwrap_or_default()
    }
}

fn count(len: usize) -> Result<u32> {
    u32::try_from(len).map_err(|_| Error::from_hresult(E_INVALIDARG))
}
