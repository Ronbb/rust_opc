//! Safe OPC Data Access client facade.

mod group;
mod properties;
mod types;

use std::marker::PhantomData;

use opc_classic_types::{ClassContext, ComObject, Error, ErrorCode, Guid, Result};
use opc_classic_utils::{Cleanup, CoTaskMemArrayOut, CoTaskMemOut, ComApartment};
use opc_comn_bindings::client::CommonClient;
use windows::Win32::System::Com::{CLSIDFromProgID, CoCreateInstance, CoTaskMemFree};
use windows_core::{IUnknown, Interface, PCWSTR};

use crate::abi::{
    class_context_to_abi, error_code_to_abi, from_abi_error, guid_to_abi, interface_from_object,
    object_from_interface, timestamp_from_abi,
};
use crate::{IOPCItemMgt, IOPCServer, tagOPCSERVERSTATUS};

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
        apartment: &'apartment ComApartment,
        class_id: &Guid,
        context: ClassContext,
    ) -> Result<Self> {
        let class_id = guid_to_abi(class_id);
        let context = class_context_to_abi(context);
        let inner = unsafe { CoCreateInstance(&class_id, None::<&IUnknown>, context) }
            .map_err(from_abi_error)?;
        Ok(Self::from_interface(apartment, inner))
    }

    pub fn from_object(apartment: &'apartment ComApartment, object: &ComObject) -> Result<Self> {
        let inner = interface_from_object(object)?;
        Ok(Self::from_interface(apartment, inner))
    }

    pub(crate) fn from_interface(apartment: &'apartment ComApartment, inner: IOPCServer) -> Self {
        let _ = apartment;
        Self {
            inner,
            _apartment: PhantomData,
        }
    }

    pub fn common(&self) -> Result<CommonClient> {
        CommonClient::from_object(&self.object())
    }

    pub fn properties(&self) -> Result<ItemPropertiesClient> {
        ItemPropertiesClient::from_object(&self.object())
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
                PCWSTR(name.as_ptr()),
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
        }
        .map_err(from_abi_error)?;
        let object =
            object.ok_or_else(|| Error::unexpected("AddGroup returned no group object"))?;
        Ok(DaGroup::new(
            self.inner.clone(),
            object,
            server_handle,
            true,
            revised_update_rate,
        ))
    }

    pub fn status(&self) -> Result<ServerStatus> {
        let mut status = CoTaskMemArrayOut::new(1, ServerStatusCleanup);
        let call = unsafe {
            (Interface::vtable(&self.inner).GetStatus)(
                Interface::as_raw(&self.inner),
                status.as_mut_ptr(),
            )
        };
        let status = unsafe { status.into_array() };
        call.ok().map_err(from_abi_error)?;
        let status = status?;
        let raw = &status.as_slice()[0];
        let vendor_info = if raw.szVendorInfo.is_null() {
            String::new()
        } else {
            unsafe { raw.szVendorInfo.to_string() }
                .map_err(|_| Error::invalid_argument("invalid UTF-16 vendor string"))?
        };
        Ok(ServerStatus {
            start_time: timestamp_from_abi(raw.ftStartTime),
            current_time: timestamp_from_abi(raw.ftCurrentTime),
            last_update_time: timestamp_from_abi(raw.ftLastUpdateTime),
            state: ServerState::from_raw(raw.dwServerState.0),
            group_count: raw.dwGroupCount,
            bandwidth: raw.dwBandWidth,
            version: (raw.wMajorVersion, raw.wMinorVersion, raw.wBuildNumber),
            vendor_info,
        })
    }

    pub fn error_string(&self, error: ErrorCode, locale: u32) -> Result<String> {
        let mut value = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            (Interface::vtable(&self.inner).GetErrorString)(
                Interface::as_raw(&self.inner),
                error_code_to_abi(error),
                locale,
                value.as_mut_ptr().cast(),
            )
        };
        let value = unsafe { value.into_pwstr() };
        call.ok().map_err(from_abi_error)?;
        Ok(value.to_string_lossy())
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }
}
