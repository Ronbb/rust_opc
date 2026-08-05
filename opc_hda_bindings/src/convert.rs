use core::mem::ManuallyDrop;

use opc_classic_types::{ComObject, Error, ErrorCode, Guid, Result, Timestamp, Value};
use windows::Win32::Foundation::FILETIME;
use windows::Win32::System::Variant::{
    VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8,
    VT_NULL, VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8,
};
use windows_core::{BSTR, Interface};

pub(crate) fn from_abi_error(error: windows_core::Error) -> Error {
    Error::from_code(ErrorCode::from_raw(error.code().0)).with_message(error.message())
}

pub(crate) fn to_abi_error(error: Error) -> windows_core::Error {
    windows_core::Error::from_hresult(windows_core::HRESULT(error.code().raw()))
}

pub(crate) fn guid_to_abi(value: &Guid) -> windows_core::GUID {
    windows_core::GUID::from_values(value.data1, value.data2, value.data3, value.data4)
}

pub(crate) fn timestamp_from_abi(value: FILETIME) -> Timestamp {
    Timestamp::from_ticks((u64::from(value.dwHighDateTime) << 32) | u64::from(value.dwLowDateTime))
}

pub(crate) fn timestamp_to_abi(value: Timestamp) -> FILETIME {
    FILETIME {
        dwLowDateTime: value.ticks() as u32,
        dwHighDateTime: (value.ticks() >> 32) as u32,
    }
}

pub(crate) fn object_from_interface<T: Interface>(value: T) -> ComObject {
    unsafe { ComObject::from_raw_owned(value.into_raw()).expect("COM interfaces are non-null") }
}

pub(crate) fn interface_from_object<T: Interface>(value: &ComObject) -> Result<T> {
    let iid = T::IID;
    let iid = Guid::new(iid.data1, iid.data2, iid.data3, iid.data4);
    let requested = value.query_interface(&iid)?;
    Ok(unsafe { T::from_raw(requested.into_raw()) })
}

pub(crate) fn value_to_abi(value: &Value) -> Result<VARIANT> {
    let variant = match value {
        Value::Empty => VARIANT::default(),
        Value::Null => scalar_variant(VT_NULL.0, 0),
        Value::Bool(value) => VARIANT::from(*value),
        Value::I8(value) => VARIANT::from(*value),
        Value::U8(value) => VARIANT::from(*value),
        Value::I16(value) => VARIANT::from(*value),
        Value::U16(value) => VARIANT::from(*value),
        Value::I32(value) => VARIANT::from(*value),
        Value::U32(value) => VARIANT::from(*value),
        Value::I64(value) => VARIANT::from(*value),
        Value::U64(value) => VARIANT::from(*value),
        Value::F32(value) => VARIANT::from(*value),
        Value::F64(value) => VARIANT::from(*value),
        Value::Currency(value) => scalar_variant(VT_CY.0, *value as u64),
        Value::Date(value) => scalar_variant(VT_DATE.0, value.to_bits()),
        Value::String(value) => VARIANT::from(value.as_str()),
        Value::Error(value) => scalar_variant(VT_ERROR.0, value.raw() as u32 as u64),
        Value::Bytes(_) | Value::Array(_) => {
            return Err(Error::not_implemented(
                "Automation array conversion is not implemented",
            ));
        }
        _ => return Err(Error::not_implemented("unsupported Automation value")),
    };
    Ok(variant)
}

pub(crate) fn value_from_abi(value: &VARIANT) -> Result<Value> {
    let inner = unsafe { &*value.Anonymous.Anonymous };
    let data = &inner.Anonymous;
    match inner.vt {
        VT_EMPTY => Ok(Value::Empty),
        VT_NULL => Ok(Value::Null),
        VT_BOOL => Ok(Value::Bool(unsafe { data.boolVal.0 != 0 })),
        VT_I1 => Ok(Value::I8(unsafe { data.cVal })),
        VT_UI1 => Ok(Value::U8(unsafe { data.bVal })),
        VT_I2 => Ok(Value::I16(unsafe { data.iVal })),
        VT_UI2 => Ok(Value::U16(unsafe { data.uiVal })),
        VT_I4 => Ok(Value::I32(unsafe { data.lVal })),
        VT_UI4 => Ok(Value::U32(unsafe { data.ulVal })),
        VT_I8 => Ok(Value::I64(unsafe { data.llVal })),
        VT_UI8 => Ok(Value::U64(unsafe { data.ullVal })),
        VT_R4 => Ok(Value::F32(unsafe { data.fltVal })),
        VT_R8 => Ok(Value::F64(unsafe { data.dblVal })),
        VT_CY => Ok(Value::Currency(unsafe { data.cyVal.int64 })),
        VT_DATE => Ok(Value::Date(unsafe { data.date })),
        VT_ERROR => Ok(Value::Error(ErrorCode::from_raw(unsafe { data.scode }))),
        VT_BSTR => {
            let bstr = unsafe { &*((&data.bstrVal as *const ManuallyDrop<BSTR>).cast::<BSTR>()) };
            String::try_from(bstr)
                .map(Value::String)
                .map_err(|_| Error::invalid_argument("Automation string is invalid UTF-16"))
        }
        other => Err(Error::new(
            opc_classic_types::ErrorKind::Automation,
            ErrorCode::TYPE_MISMATCH,
            format!("unsupported Automation value type {}", other.0),
        )),
    }
}

fn scalar_variant(value_type: u16, bits: u64) -> VARIANT {
    use windows::Win32::System::Variant::{VARENUM, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0};

    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt: VARENUM(value_type),
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: VARIANT_0_0_0 { ullVal: bits },
            }),
        },
    }
}
