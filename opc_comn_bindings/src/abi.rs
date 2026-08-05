//! Private conversions between the generated Windows ABI and the public,
//! Windows-independent OPC Classic types.

use opc_classic_types::{ComObject, Error, ErrorCode, Guid, Result};
use windows_core::{Error as AbiError, GUID, Interface};

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

pub(crate) const fn from_abi_guid(value: GUID) -> Guid {
    Guid::new(value.data1, value.data2, value.data3, value.data4)
}

pub(crate) const fn to_abi_guid(value: Guid) -> GUID {
    GUID::from_values(value.data1, value.data2, value.data3, value.data4)
}

pub(crate) fn interface_from_object<T: Interface>(object: &ComObject) -> Result<T> {
    let interface_id = from_abi_guid(T::IID);
    let interface = object.query_interface(&interface_id)?;
    // SAFETY: `query_interface` returned one owned reference for `T::IID`.
    Ok(unsafe { T::from_raw(interface.into_raw()) })
}

/// Adopts an interface pointer written by a raw COM call before its status is
/// inspected. Keeping the resulting `Result<T>` alive ensures that a non-null
/// pointer is released even when a broken server returns a failure status after
/// writing the out parameter.
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
    // Cloning supplies the owned reference transferred to `ComObject`.
    let raw = interface.clone().into_raw();
    // SAFETY: every generated interface is a non-null COM interface pointer.
    unsafe { ComObject::from_raw_owned(raw) }
        .expect("a generated COM interface must contain a non-null pointer")
}
