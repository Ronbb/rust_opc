//! Safe OPC Alarms & Events client facade.

use std::marker::PhantomData;

use opc_classic_types::{ClassContext, ComObject, Error, Guid, Result, Timestamp};
use opc_classic_utils::{
    CoTaskMemArrayOut, CoTaskMemOut, ComApartment, FreePwstrElements, NoCleanup, WideCString,
};
use opc_comn_bindings::client::CommonClient;
use windows_core::{BOOL, Interface, PCWSTR, PWSTR, Result as AbiResult};

use crate::abi::{
    StatusCleanup, class_id_from_prog_id, create_instance, from_abi_error, from_abi_timestamp,
    interface_from_object, interface_from_raw_owned, object_from_interface,
};
use crate::{
    __MIDL___MIDL_itf_opc_ae_0000_0001_0001, AeServerState, BrowseDirection, IOPCEventAreaBrowser,
    IOPCEventServer, IOPCEventSubscriptionMgt, IOPCEventSubscriptionMgt2,
};

#[derive(Clone, Debug)]
pub struct AeServerStatus {
    pub start_time: Timestamp,
    pub current_time: Timestamp,
    pub last_update_time: Timestamp,
    pub state: AeServerState,
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
    pub fn from_object(object: &ComObject) -> Result<Self> {
        Ok(Self {
            inner: interface_from_object(object)?,
            revised_buffer_time: 0,
            revised_max_size: 0,
        })
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
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
        let area_strings = areas
            .iter()
            .map(|area| wide(area))
            .collect::<Result<Vec<_>>>()?;
        let source_strings = sources
            .iter()
            .map(|source| wide(source))
            .collect::<Result<Vec<_>>>()?;
        let area_pointers: Vec<_> = area_strings
            .iter()
            .map(|value| PCWSTR(value.as_ptr()))
            .collect();
        let source_pointers: Vec<_> = source_strings
            .iter()
            .map(|value| PCWSTR(value.as_ptr()))
            .collect();
        unsafe {
            self.inner.SetFilter(
                event_type,
                categories,
                low_severity,
                high_severity,
                &area_pointers,
                &source_pointers,
            )
        }
        .map_err(from_abi_error)
    }

    pub fn filter(&self) -> Result<EventFilter> {
        let mut event_type = 0;
        let mut category_count = 0;
        let mut categories = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut low_severity = 0;
        let mut high_severity = 0;
        let mut area_count = 0;
        let mut areas = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut source_count = 0;
        let mut sources = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let call = unsafe {
            self.inner.GetFilter(
                &mut event_type,
                &mut category_count,
                categories.as_mut_ptr(),
                &mut low_severity,
                &mut high_severity,
                &mut area_count,
                areas.as_mut_ptr().cast::<*mut PWSTR>(),
                &mut source_count,
                sources.as_mut_ptr().cast::<*mut PWSTR>(),
            )
        };
        unsafe {
            categories.set_len(category_count as usize);
            areas.set_len(area_count as usize);
            sources.set_len(source_count as usize);
        }
        let categories = unsafe { categories.into_array() };
        let areas = unsafe { areas.into_array() };
        let sources = unsafe { sources.into_array() };
        call.map_err(from_abi_error)?;
        let categories = categories?;
        let areas = areas?;
        let sources = sources?;
        Ok(EventFilter {
            event_type,
            categories: categories.as_slice().to_vec(),
            low_severity,
            high_severity,
            areas: areas
                .as_slice()
                .iter()
                .map(|value| pwstr_string(PWSTR(*value)))
                .collect(),
            sources: sources
                .as_slice()
                .iter()
                .map(|value| pwstr_string(PWSTR(*value)))
                .collect(),
        })
    }

    pub fn select_returned_attributes(&self, category: u32, attributes: &[u32]) -> Result<()> {
        unsafe { self.inner.SelectReturnedAttributes(category, attributes) }.map_err(from_abi_error)
    }

    pub fn returned_attributes(&self, category: u32) -> Result<Vec<u32>> {
        let mut count = 0;
        let mut values = CoTaskMemArrayOut::new(0, NoCleanup);
        let call = unsafe {
            self.inner
                .GetReturnedAttributes(category, &mut count, values.as_mut_ptr())
        };
        unsafe { values.set_len(count as usize) };
        let values = unsafe { values.into_array() };
        call.map_err(from_abi_error)?;
        let values = values?;
        Ok(values.as_slice().to_vec())
    }

    pub fn refresh(&self, connection: u32) -> Result<()> {
        unsafe { self.inner.Refresh(connection) }.map_err(from_abi_error)
    }

    pub fn cancel_refresh(&self, connection: u32) -> Result<()> {
        unsafe { self.inner.CancelRefresh(connection) }.map_err(from_abi_error)
    }

    pub fn state(&self) -> Result<SubscriptionState> {
        let mut active = BOOL(0);
        let mut buffer_time = 0;
        let mut max_size = 0;
        let mut client_handle = 0;
        unsafe {
            self.inner.GetState(
                &mut active,
                &mut buffer_time,
                &mut max_size,
                &mut client_handle,
            )
        }
        .map_err(from_abi_error)?;
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
            )
        }
        .map_err(from_abi_error)?;
        Ok((revised_buffer_time, revised_max_size))
    }

    pub fn keep_alive(&self) -> Result<KeepAlive> {
        let inner = interface_from_object(&self.object())?;
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
    pub fn from_object(object: &ComObject) -> Result<Self> {
        Ok(Self {
            inner: interface_from_object(object)?,
        })
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn set(&self, requested: u32) -> Result<u32> {
        unsafe { self.inner.SetKeepAlive(requested) }.map_err(from_abi_error)
    }

    pub fn get(&self) -> Result<u32> {
        unsafe { self.inner.GetKeepAlive() }.map_err(from_abi_error)
    }
}

#[derive(Clone)]
pub struct EventAreaBrowser {
    inner: IOPCEventAreaBrowser,
}

impl EventAreaBrowser {
    pub fn from_object(object: &ComObject) -> Result<Self> {
        Ok(Self {
            inner: interface_from_object(object)?,
        })
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn browse_position(&self, direction: BrowseDirection, name: &str) -> Result<()> {
        let name = wide(name)?;
        let direction = __MIDL___MIDL_itf_opc_ae_0000_0001_0001(direction.raw());
        unsafe {
            self.inner
                .ChangeBrowsePosition(direction, PCWSTR(name.as_ptr()))
        }
        .map_err(from_abi_error)
    }

    pub fn qualified_area_name(&self, name: &str) -> Result<String> {
        let name = wide(name)?;
        let mut value = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            (Interface::vtable(&self.inner).GetQualifiedAreaName)(
                Interface::as_raw(&self.inner),
                PCWSTR(name.as_ptr()),
                value.as_mut_ptr().cast(),
            )
        };
        let value = unsafe { value.into_pwstr() };
        call.ok().map_err(from_abi_error)?;
        Ok(value.to_string_lossy())
    }

    pub fn qualified_source_name(&self, name: &str) -> Result<String> {
        let name = wide(name)?;
        let mut value = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            (Interface::vtable(&self.inner).GetQualifiedSourceName)(
                Interface::as_raw(&self.inner),
                PCWSTR(name.as_ptr()),
                value.as_mut_ptr().cast(),
            )
        };
        let value = unsafe { value.into_pwstr() };
        call.ok().map_err(from_abi_error)?;
        Ok(value.to_string_lossy())
    }
}

pub struct AeClient<'apartment> {
    inner: IOPCEventServer,
    _apartment: PhantomData<&'apartment ComApartment>,
}

impl<'apartment> AeClient<'apartment> {
    pub fn connect_prog_id(apartment: &'apartment ComApartment, prog_id: &str) -> Result<Self> {
        let class_id = class_id_from_prog_id(prog_id)?;
        Self::connect_class_id(apartment, &class_id, ClassContext::ALL)
    }

    pub fn connect_class_id(
        _apartment: &'apartment ComApartment,
        class_id: &Guid,
        context: ClassContext,
    ) -> Result<Self> {
        let inner = create_instance(*class_id, context)?;
        Ok(Self {
            inner,
            _apartment: PhantomData,
        })
    }

    pub fn from_object(_apartment: &'apartment ComApartment, object: &ComObject) -> Result<Self> {
        Ok(Self {
            inner: interface_from_object(object)?,
            _apartment: PhantomData,
        })
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn common(&self) -> Result<CommonClient> {
        CommonClient::from_object(&self.object())
    }

    pub fn status(&self) -> Result<AeServerStatus> {
        let mut status = CoTaskMemArrayOut::new(1, StatusCleanup);
        let call = unsafe {
            (Interface::vtable(&self.inner).GetStatus)(
                Interface::as_raw(&self.inner),
                status.as_mut_ptr(),
            )
        };
        let status = unsafe { status.into_array() };
        call.ok().map_err(from_abi_error)?;
        let status = status?;
        let value = &status.as_slice()[0];
        Ok(AeServerStatus {
            start_time: from_abi_timestamp(value.ftStartTime),
            current_time: from_abi_timestamp(value.ftCurrentTime),
            last_update_time: from_abi_timestamp(value.ftLastUpdateTime),
            state: AeServerState::from_raw(value.dwServerState.0)?,
            version: (value.wMajorVersion, value.wMinorVersion, value.wBuildNumber),
            vendor_info: pwstr_string(value.szVendorInfo),
        })
    }

    pub fn available_filters(&self) -> Result<u32> {
        unsafe { self.inner.QueryAvailableFilters() }.map_err(from_abi_error)
    }

    pub fn event_categories(&self, event_type: u32) -> Result<Vec<EventCategory>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut descriptions = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let call = unsafe {
            self.inner.QueryEventCategories(
                event_type,
                &mut count,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr().cast::<*mut PWSTR>(),
            )
        };
        let len = count as usize;
        unsafe {
            ids.set_len(len);
            descriptions.set_len(len);
        }
        let ids = unsafe { ids.into_array() };
        let descriptions = unsafe { descriptions.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let descriptions = descriptions?;
        Ok(ids
            .as_slice()
            .iter()
            .zip(descriptions.as_slice())
            .map(|(id, description)| EventCategory {
                id: *id,
                description: pwstr_string(PWSTR(*description)),
            })
            .collect())
    }

    pub fn condition_names(&self, event_category: u32) -> Result<Vec<String>> {
        self.query_string_array(|count, output| unsafe {
            self.inner
                .QueryConditionNames(event_category, count, output)
        })
    }

    pub fn subcondition_names(&self, condition: &str) -> Result<Vec<String>> {
        let condition = wide(condition)?;
        self.query_string_array(|count, output| unsafe {
            self.inner
                .QuerySubConditionNames(PCWSTR(condition.as_ptr()), count, output)
        })
    }

    pub fn source_conditions(&self, source: &str) -> Result<Vec<String>> {
        let source = wide(source)?;
        self.query_string_array(|count, output| unsafe {
            self.inner
                .QuerySourceConditions(PCWSTR(source.as_ptr()), count, output)
        })
    }

    pub fn event_attributes(&self, event_category: u32) -> Result<Vec<EventAttribute>> {
        let mut count = 0u32;
        let mut ids = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut descriptions = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut types = CoTaskMemArrayOut::new(0, NoCleanup);
        let call = unsafe {
            self.inner.QueryEventAttributes(
                event_category,
                &mut count,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr().cast::<*mut PWSTR>(),
                types.as_mut_ptr(),
            )
        };
        let len = count as usize;
        unsafe {
            ids.set_len(len);
            descriptions.set_len(len);
            types.set_len(len);
        }
        let ids = unsafe { ids.into_array() };
        let descriptions = unsafe { descriptions.into_array() };
        let types = unsafe { types.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let descriptions = descriptions?;
        let types = types?;
        Ok((0..len)
            .map(|index| EventAttribute {
                id: ids.as_slice()[index],
                description: pwstr_string(PWSTR(descriptions.as_slice()[index])),
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
        }
        .map_err(from_abi_error)?;
        let object =
            object.ok_or_else(|| Error::unexpected("AE server returned a null subscription"))?;
        let identity = object_from_interface(&object);
        Ok(EventSubscription {
            inner: interface_from_object(&identity)?,
            revised_buffer_time,
            revised_max_size,
        })
    }

    pub fn area_browser(&self) -> Result<EventAreaBrowser> {
        let mut raw = core::ptr::null_mut();
        let call = unsafe {
            (Interface::vtable(&self.inner).CreateAreaBrowser)(
                Interface::as_raw(&self.inner),
                &IOPCEventAreaBrowser::IID,
                &mut raw,
            )
        };
        let browser = unsafe { interface_from_raw_owned(raw) };
        call.ok().map_err(from_abi_error)?;
        Ok(EventAreaBrowser { inner: browser? })
    }

    fn query_string_array(
        &self,
        call: impl FnOnce(*mut u32, *mut *mut PWSTR) -> AbiResult<()>,
    ) -> Result<Vec<String>> {
        let mut count = 0u32;
        let mut names = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let result = call(&mut count, names.as_mut_ptr().cast::<*mut PWSTR>());
        unsafe { names.set_len(count as usize) };
        let names = unsafe { names.into_array() };
        result.map_err(from_abi_error)?;
        let names = names?;
        Ok(names
            .as_slice()
            .iter()
            .map(|name| pwstr_string(PWSTR(*name)))
            .collect())
    }

    fn set_condition_state(&self, names: &[&str], enabled: bool, area: bool) -> Result<()> {
        let wide_names = names
            .iter()
            .map(|name| wide(name))
            .collect::<Result<Vec<_>>>()?;
        let pointers: Vec<PCWSTR> = wide_names
            .iter()
            .map(|value| PCWSTR(value.as_ptr()))
            .collect();
        if area {
            if enabled {
                unsafe { self.inner.EnableConditionByArea(&pointers) }.map_err(from_abi_error)
            } else {
                unsafe { self.inner.DisableConditionByArea(&pointers) }.map_err(from_abi_error)
            }
        } else if enabled {
            unsafe { self.inner.EnableConditionBySource(&pointers) }.map_err(from_abi_error)
        } else {
            unsafe { self.inner.DisableConditionBySource(&pointers) }.map_err(from_abi_error)
        }
    }
}

fn wide(value: &str) -> Result<WideCString> {
    WideCString::try_from(value)
        .map_err(|_| Error::invalid_argument("string contains an interior NUL"))
}

fn pwstr_string(value: PWSTR) -> String {
    if value.is_null() {
        String::new()
    } else {
        String::from_utf16_lossy(unsafe { value.as_wide() })
    }
}
