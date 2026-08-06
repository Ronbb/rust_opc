//! Private conversions between the generated Windows ABI and the public API.

use opc_classic_types::{ClassContext, ComObject, Error, ErrorCode, Guid, Result, Timestamp};
use opc_classic_utils::Cleanup;
use windows::Win32::Foundation::FILETIME;
use windows::Win32::System::Com::{CLSCTX, CLSIDFromProgID, CoCreateInstance, CoTaskMemFree};
use windows_core::{Error as AbiError, GUID, IUnknown, Interface, PCWSTR, PWSTR};

pub(crate) struct StatusCleanup;

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

pub(crate) fn from_abi_error(error: AbiError) -> Error {
    Error::from_code(ErrorCode::from_raw(error.code().0)).with_message(error.message())
}

pub(crate) fn to_abi_error(error: Error) -> AbiError {
    let status = error.boundary_status();
    let status = windows_core::HRESULT(status.raw());
    if status.0 == ErrorCode::PARTIAL_SUCCESS.raw() {
        AbiError::from_hresult(status)
    } else {
        AbiError::new(status, error.message())
    }
}

pub(crate) const fn to_abi_guid(value: Guid) -> GUID {
    GUID::from_values(value.data1, value.data2, value.data3, value.data4)
}

pub(crate) const fn from_abi_timestamp(value: FILETIME) -> Timestamp {
    Timestamp::from_ticks(((value.dwHighDateTime as u64) << 32) | value.dwLowDateTime as u64)
}

pub(crate) const fn to_abi_timestamp(value: Timestamp) -> FILETIME {
    FILETIME {
        dwLowDateTime: value.ticks() as u32,
        dwHighDateTime: (value.ticks() >> 32) as u32,
    }
}

pub(crate) fn interface_from_object<T: Interface>(object: &ComObject) -> Result<T> {
    let iid = Guid::new(T::IID.data1, T::IID.data2, T::IID.data3, T::IID.data4);
    let interface = object.query_interface(&iid)?;
    // SAFETY: QueryInterface returned one owned reference for `T::IID`.
    Ok(unsafe { T::from_raw(interface.into_raw()) })
}

pub(crate) unsafe fn interface_from_raw_owned<T: Interface>(
    ptr: *mut core::ffi::c_void,
) -> Result<T> {
    if ptr.is_null() {
        Err(Error::null_pointer("COM call returned a null interface"))
    } else {
        Ok(unsafe { T::from_raw(ptr) })
    }
}

pub(crate) fn object_from_interface<T: Interface>(interface: &T) -> ComObject {
    let raw = interface.clone().into_raw();
    // SAFETY: generated COM interfaces always contain a non-null owned pointer.
    unsafe { ComObject::from_raw_owned(raw) }
        .expect("a generated COM interface must contain a non-null pointer")
}

pub(crate) fn create_instance<T: Interface>(class_id: Guid, context: ClassContext) -> Result<T> {
    let class_id = to_abi_guid(class_id);
    unsafe { CoCreateInstance(&class_id, None::<&IUnknown>, CLSCTX(context.bits())) }
        .map_err(from_abi_error)
}

pub(crate) fn class_id_from_prog_id(prog_id: &str) -> Result<Guid> {
    let prog_id = opc_classic_utils::WideCString::try_from(prog_id)
        .map_err(|_| Error::invalid_argument("ProgID contains an interior NUL"))?;
    let value = unsafe { CLSIDFromProgID(PCWSTR(prog_id.as_ptr())) }.map_err(from_abi_error)?;
    Ok(Guid::new(
        value.data1,
        value.data2,
        value.data3,
        value.data4,
    ))
}
