//! Safe OPC Alarms & Events client facade.

use std::marker::PhantomData;

use opc_classic_utils::{
    Cleanup, CoTaskMemArray, CoTaskMemOut, ComApartment, FreePwstrElements, NoCleanup, OwnedPwstr,
    WideCString,
};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::Foundation::{E_INVALIDARG, E_UNEXPECTED, FILETIME};
use windows::Win32::System::Com::{
    CLSCTX, CLSCTX_ALL, CLSIDFromProgID, CoCreateInstance, CoTaskMemFree,
};
use windows_core::{Error, GUID, IUnknown, Interface, PCWSTR, Result};

use crate::{
    __MIDL___MIDL_itf_opc_ae_0000_0001_0001, __MIDL___MIDL_itf_opc_ae_0000_0001_0003,
    __MIDL___MIDL_itf_opc_ae_0000_0001_0005, IOPCEventAreaBrowser, IOPCEventServer,
    IOPCEventSubscriptionMgt, IOPCEventSubscriptionMgt2,
};

#[derive(Clone, Debug)]
pub struct AeServerStatus {
    pub start_time: FILETIME,
    pub current_time: FILETIME,
    pub last_update_time: FILETIME,
    pub state: __MIDL___MIDL_itf_opc_ae_0000_0001_0003,
    pub version: (u16, u16, u16),
    pub vendor_info: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EventCategory {
    pub id: u32,
    pub description: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EventAttribute {
    pub id: u32,
    pub description: String,
    pub data_type: u16,
}

#[derive(Clone, Debug)]
pub struct EventSubscriptionOptions {
    pub active: bool,
    pub buffer_time: u32,
    pub max_size: u32,
    pub client_handle: u32,
}

#[derive(Clone)]
pub struct EventSubscription {
    inner: IOPCEventSubscriptionMgt,
    pub revised_buffer_time: u32,
    pub revised_max_size: u32,
}

impl EventSubscription {
    pub fn as_raw(&self) -> &IOPCEventSubscriptionMgt {
        &self.inner
    }

    pub fn set_filter(
        &self,
        event_type: u32,
        categories: &[u32],
        low_severity: u32,
        high_severity: u32,
        areas: &[&str],
        sources: &[&str],
    ) -> Result<()> {
        let areas = areas
            .iter()
            .map(|area| wide(area))
            .collect::<Result<Vec<_>>>()?;
        let sources = sources
            .iter()
            .map(|source| wide(source))
            .collect::<Result<Vec<_>>>()?;
        let areas: Vec<_> = areas.iter().map(WideCString::as_pcwstr).collect();
        let sources: Vec<_> = sources.iter().map(WideCString::as_pcwstr).collect();
        unsafe {
            self.inner.SetFilter(
                event_type,
                categories,
                low_severity,
                high_severity,
                &areas,
                &sources,
            )
        }
    }

    pub fn filter(&self) -> Result<EventFilter> {
        let mut event_type = 0;
        let mut category_count = 0;
        let mut categories = CoTaskMemOut::<u32>::new();
        let mut low_severity = 0;
        let mut high_severity = 0;
        let mut area_count = 0;
        let mut areas = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut source_count = 0;
        let mut sources = CoTaskMemOut::<windows_core::PWSTR>::new();
        let call = unsafe {
            self.inner.GetFilter(
                &mut event_type,
                &mut category_count,
                categories.as_mut_ptr(),
                &mut low_severity,
                &mut high_severity,
                &mut area_count,
                areas.as_mut_ptr(),
                &mut source_count,
                sources.as_mut_ptr(),
            )
        };
        let categories = unsafe { categories.into_array(category_count as usize, NoCleanup) }?;
        let areas = unsafe { areas.into_array(area_count as usize, FreePwstrElements) }?;
        let sources = unsafe { sources.into_array(source_count as usize, FreePwstrElements) }?;
        call?;
        Ok(EventFilter {
            event_type,
            categories: categories.as_slice().to_vec(),
            low_severity,
            high_severity,
            areas: areas
                .as_slice()
                .iter()
                .map(|value| pwstr_string(*value))
                .collect(),
            sources: sources
                .as_slice()
                .iter()
                .map(|value| pwstr_string(*value))
                .collect(),
        })
    }

    pub fn select_returned_attributes(&self, category: u32, attributes: &[u32]) -> Result<()> {
        unsafe { self.inner.SelectReturnedAttributes(category, attributes) }
    }

    pub fn returned_attributes(&self, category: u32) -> Result<Vec<u32>> {
        let mut count = 0;
        let mut values = CoTaskMemOut::<u32>::new();
        let call = unsafe {
            self.inner
                .GetReturnedAttributes(category, &mut count, values.as_mut_ptr())
        };
        let values = unsafe { values.into_array(count as usize, NoCleanup) }?;
        call?;
        Ok(values.as_slice().to_vec())
    }

    pub fn refresh(&self, connection: u32) -> Result<()> {
        unsafe { self.inner.Refresh(connection) }
    }

    pub fn cancel_refresh(&self, connection: u32) -> Result<()> {
        unsafe { self.inner.CancelRefresh(connection) }
    }

    pub fn state(&self) -> Result<SubscriptionState> {
        let mut active = windows_core::BOOL(0);
        let mut buffer_time = 0;
        let mut max_size = 0;
        let mut client_handle = 0;
        unsafe {
            self.inner.GetState(
                &mut active,
                &mut buffer_time,
                &mut max_size,
                &mut client_handle,
            )?
        };
        Ok(SubscriptionState {
            active: active.as_bool(),
            buffer_time,
            max_size,
            client_handle,
        })
    }

    pub fn set_state(&self, state: SubscriptionState) -> Result<(u32, u32)> {
        let active = state.active.into();
        let mut revised_buffer_time = 0;
        let mut revised_max_size = 0;
        unsafe {
            self.inner.SetState(
                &active,
                &state.buffer_time,
                &state.max_size,
                state.client_handle,
                &mut revised_buffer_time,
                &mut revised_max_size,
            )?
        };
        Ok((revised_buffer_time, revised_max_size))
    }

    pub fn keep_alive(&self) -> Result<KeepAlive> {
        let inner: IOPCEventSubscriptionMgt2 = self.inner.cast()?;
        Ok(KeepAlive { inner })
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EventFilter {
    pub event_type: u32,
    pub categories: Vec<u32>,
    pub low_severity: u32,
    pub high_severity: u32,
    pub areas: Vec<String>,
    pub sources: Vec<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SubscriptionState {
    pub active: bool,
    pub buffer_time: u32,
    pub max_size: u32,
    pub client_handle: u32,
}

#[derive(Clone)]
pub struct KeepAlive {
    inner: IOPCEventSubscriptionMgt2,
}

impl KeepAlive {
    pub fn set(&self, requested: u32) -> Result<u32> {
        unsafe { self.inner.SetKeepAlive(requested) }
    }

    pub fn get(&self) -> Result<u32> {
        unsafe { self.inner.GetKeepAlive() }
    }
}

#[derive(Clone)]
pub struct EventAreaBrowser {
    inner: IOPCEventAreaBrowser,
}

impl EventAreaBrowser {
    pub fn browse_position(
        &self,
        direction: __MIDL___MIDL_itf_opc_ae_0000_0001_0001,
        name: &str,
    ) -> Result<()> {
        let name = wide(name)?;
        unsafe { self.inner.ChangeBrowsePosition(direction, name.as_pcwstr()) }
    }

    pub fn qualified_area_name(&self, name: &str) -> Result<String> {
        let name = wide(name)?;
        let value = unsafe { self.inner.GetQualifiedAreaName(name.as_pcwstr()) }?;
        Ok(unsafe { OwnedPwstr::from_raw(value.0) }.to_string_lossy())
    }

    pub fn qualified_source_name(&self, name: &str) -> Result<String> {
        let name = wide(name)?;
        let value = unsafe { self.inner.GetQualifiedSourceName(name.as_pcwstr()) }?;
        Ok(unsafe { OwnedPwstr::from_raw(value.0) }.to_string_lossy())
    }
}

pub struct AeClient<'apartment> {
    inner: IOPCEventServer,
    _apartment: PhantomData<&'apartment ComApartment>,
}

struct StatusCleanup;

// SAFETY: The status owns one task-allocated vendor string.
unsafe impl Cleanup<__MIDL___MIDL_itf_opc_ae_0000_0001_0005> for StatusCleanup {
    unsafe fn cleanup(
        &mut self,
        ptr: *mut __MIDL___MIDL_itf_opc_ae_0000_0001_0005,
        initialized: usize,
    ) {
        for index in 0..initialized {
            let value = unsafe { &mut *ptr.add(index) };
            if !value.szVendorInfo.is_null() {
                unsafe { CoTaskMemFree(Some(value.szVendorInfo.0.cast())) };
                value.szVendorInfo = windows_core::PWSTR::null();
            }
        }
    }
}

impl<'apartment> AeClient<'apartment> {
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

    pub fn from_interface(apartment: &'apartment ComApartment, inner: IOPCEventServer) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn common(&self) -> Result<CommonClient> {
        self.inner.cast().map(CommonClient::new)
    }

    pub fn status(&self) -> Result<AeServerStatus> {
        let ptr = unsafe { self.inner.GetStatus() }?;
        let status = unsafe { CoTaskMemArray::from_raw_parts(ptr, 1, StatusCleanup) }?;
        let value = &status.as_slice()[0];
        Ok(AeServerStatus {
            start_time: value.ftStartTime,
            current_time: value.ftCurrentTime,
            last_update_time: value.ftLastUpdateTime,
            state: value.dwServerState,
            version: (value.wMajorVersion, value.wMinorVersion, value.wBuildNumber),
            vendor_info: pwstr_string(value.szVendorInfo),
        })
    }

    pub fn available_filters(&self) -> Result<u32> {
        unsafe { self.inner.QueryAvailableFilters() }
    }

    pub fn event_categories(&self, event_type: u32) -> Result<Vec<EventCategory>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemOut::<u32>::new();
        let mut descriptions = CoTaskMemOut::<windows_core::PWSTR>::new();
        let call = unsafe {
            self.inner.QueryEventCategories(
                event_type,
                &mut count,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr(),
            )
        };
        let len = count as usize;
        let ids = unsafe { ids.into_array(len, NoCleanup) }?;
        let descriptions = unsafe { descriptions.into_array(len, FreePwstrElements) }?;
        call?;
        Ok(ids
            .as_slice()
            .iter()
            .zip(descriptions.as_slice())
            .map(|(id, description)| EventCategory {
                id: *id,
                description: pwstr_string(*description),
            })
            .collect())
    }

    pub fn condition_names(&self, event_category: u32) -> Result<Vec<String>> {
        let mut count = 0u32;
        let mut names = CoTaskMemOut::<windows_core::PWSTR>::new();
        let call = unsafe {
            self.inner
                .QueryConditionNames(event_category, &mut count, names.as_mut_ptr())
        };
        let names = unsafe { names.into_array(count as usize, FreePwstrElements) }?;
        call?;
        Ok(names
            .as_slice()
            .iter()
            .map(|name| pwstr_string(*name))
            .collect())
    }

    pub fn subcondition_names(&self, condition: &str) -> Result<Vec<String>> {
        self.query_string_array(|count, output| unsafe {
            self.inner
                .QuerySubConditionNames(wide(condition)?.as_pcwstr(), count, output)
        })
    }

    pub fn source_conditions(&self, source: &str) -> Result<Vec<String>> {
        self.query_string_array(|count, output| unsafe {
            self.inner
                .QuerySourceConditions(wide(source)?.as_pcwstr(), count, output)
        })
    }

    pub fn event_attributes(&self, event_category: u32) -> Result<Vec<EventAttribute>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemOut::<u32>::new();
        let mut descriptions = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut types = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            self.inner.QueryEventAttributes(
                event_category,
                &mut count,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr(),
                types.as_mut_ptr(),
            )
        };
        let len = count as usize;
        let ids = unsafe { ids.into_array(len, NoCleanup) }?;
        let descriptions = unsafe { descriptions.into_array(len, FreePwstrElements) }?;
        let types = unsafe { types.into_array(len, NoCleanup) }?;
        call?;
        Ok((0..len)
            .map(|index| EventAttribute {
                id: ids.as_slice()[index],
                description: pwstr_string(descriptions.as_slice()[index]),
                data_type: types.as_slice()[index],
            })
            .collect())
    }

    pub fn enable_areas(&self, areas: &[&str]) -> Result<()> {
        self.set_condition_state(areas, true, true)
    }

    pub fn disable_areas(&self, areas: &[&str]) -> Result<()> {
        self.set_condition_state(areas, false, true)
    }

    pub fn enable_sources(&self, sources: &[&str]) -> Result<()> {
        self.set_condition_state(sources, true, false)
    }

    pub fn disable_sources(&self, sources: &[&str]) -> Result<()> {
        self.set_condition_state(sources, false, false)
    }

    pub fn create_subscription(
        &self,
        options: EventSubscriptionOptions,
    ) -> Result<EventSubscription> {
        let mut object = None;
        let mut revised_buffer_time = 0;
        let mut revised_max_size = 0;
        unsafe {
            self.inner.CreateEventSubscription(
                options.active,
                options.buffer_time,
                options.max_size,
                options.client_handle,
                &IOPCEventSubscriptionMgt::IID,
                &mut object,
                &mut revised_buffer_time,
                &mut revised_max_size,
            )
        }?;
        let object = object.ok_or_else(|| Error::from_hresult(E_UNEXPECTED))?;
        Ok(EventSubscription {
            inner: object.cast()?,
            revised_buffer_time,
            revised_max_size,
        })
    }

    pub fn area_browser(&self) -> Result<EventAreaBrowser> {
        let object = unsafe { self.inner.CreateAreaBrowser(&IOPCEventAreaBrowser::IID) }?;
        Ok(EventAreaBrowser {
            inner: object.cast()?,
        })
    }

    pub fn as_raw(&self) -> &IOPCEventServer {
        &self.inner
    }

    fn query_string_array(
        &self,
        call: impl FnOnce(*mut u32, *mut *mut windows_core::PWSTR) -> Result<()>,
    ) -> Result<Vec<String>> {
        let mut count = 0u32;
        let mut names = CoTaskMemOut::<windows_core::PWSTR>::new();
        let result = call(&mut count, names.as_mut_ptr());
        let names = unsafe { names.into_array(count as usize, FreePwstrElements) }?;
        result?;
        Ok(names
            .as_slice()
            .iter()
            .map(|name| pwstr_string(*name))
            .collect())
    }

    fn set_condition_state(&self, names: &[&str], enabled: bool, area: bool) -> Result<()> {
        let wide_names = names
            .iter()
            .map(|name| wide(name))
            .collect::<Result<Vec<_>>>()?;
        let pointers: Vec<PCWSTR> = wide_names.iter().map(WideCString::as_pcwstr).collect();
        if area {
            if enabled {
                unsafe { self.inner.EnableConditionByArea(&pointers) }
            } else {
                unsafe { self.inner.DisableConditionByArea(&pointers) }
            }
        } else if enabled {
            unsafe { self.inner.EnableConditionBySource(&pointers) }
        } else {
            unsafe { self.inner.DisableConditionBySource(&pointers) }
        }
    }
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
