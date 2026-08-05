//! Safe OPC Data Access client facade.

mod group;
mod properties;
mod types;

use std::marker::PhantomData;

use opc_classic_utils::{Cleanup, CoTaskMemArray, ComApartment, OwnedPwstr};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::System::Com::{
    CLSCTX, CLSCTX_ALL, CLSIDFromProgID, CoCreateInstance, CoTaskMemFree,
};
use windows_core::{Error, GUID, IUnknown, Interface, Result};

use crate::{IOPCItemMgt, IOPCItemProperties, IOPCServer, tagOPCSERVERSTATUS};

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
        let class_id = unsafe { CLSIDFromProgID(prog_id.as_pcwstr()) }?;
        Self::connect_class_id(apartment, &class_id, CLSCTX_ALL)
    }

    pub fn connect_class_id(
        apartment: &'apartment ComApartment,
        class_id: &GUID,
        context: CLSCTX,
    ) -> Result<Self> {
        let inner = unsafe { CoCreateInstance(class_id, None::<&IUnknown>, context) }?;
        Ok(Self::from_interface(apartment, inner))
    }

    pub fn from_interface(apartment: &'apartment ComApartment, inner: IOPCServer) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn common(&self) -> Result<CommonClient> {
        self.inner.cast().map(CommonClient::new)
    }

    pub fn properties(&self) -> Result<ItemPropertiesClient> {
        self.inner
            .cast::<IOPCItemProperties>()
            .map(ItemPropertiesClient::new)
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
                name.as_pcwstr(),
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
        }?;
        let object =
            object.ok_or_else(|| Error::from_hresult(windows::Win32::Foundation::E_UNEXPECTED))?;
        Ok(DaGroup::new(
            self.inner.clone(),
            object,
            server_handle,
            true,
            revised_update_rate,
        ))
    }

    pub fn status(&self) -> Result<ServerStatus> {
        let ptr = unsafe { self.inner.GetStatus() }?;
        let status = unsafe { CoTaskMemArray::from_raw_parts(ptr, 1, ServerStatusCleanup) }?;
        let raw = &status.as_slice()[0];
        let vendor_info = if raw.szVendorInfo.is_null() {
            String::new()
        } else {
            unsafe { raw.szVendorInfo.to_string() }.unwrap_or_default()
        };
        Ok(ServerStatus {
            start_time: raw.ftStartTime,
            current_time: raw.ftCurrentTime,
            last_update_time: raw.ftLastUpdateTime,
            state: raw.dwServerState,
            group_count: raw.dwGroupCount,
            bandwidth: raw.dwBandWidth,
            version: (raw.wMajorVersion, raw.wMinorVersion, raw.wBuildNumber),
            vendor_info,
        })
    }

    pub fn error_string(&self, error: windows_core::HRESULT, locale: u32) -> Result<String> {
        let value = unsafe { self.inner.GetErrorString(error, locale) }?;
        Ok(unsafe { OwnedPwstr::from_raw(value.0) }.to_string_lossy())
    }

    pub fn as_raw(&self) -> &IOPCServer {
        &self.inner
    }
}
