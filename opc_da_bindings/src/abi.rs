use std::mem::ManuallyDrop;

use opc_classic_types::{
    ClassContext, ComObject, Error, ErrorCode, ErrorKind, Guid, Result, Timestamp, Value,
};
use windows::Win32::Foundation::{FILETIME, VARIANT_BOOL};
use windows::Win32::System::Com::{CLSCTX, CY};
use windows::Win32::System::Variant::{
    VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_BOOL, VT_BSTR, VT_CY, VT_DATE,
    VT_EMPTY, VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_NULL, VT_R4, VT_R8, VT_UI1, VT_UI2,
    VT_UI4, VT_UI8, VT_UINT,
};
use windows_core::{Error as AbiError, GUID, HRESULT, Interface};

pub(crate) fn from_abi_error(error: AbiError) -> Error {
    Error::from_code(ErrorCode::from_raw(error.code().0)).with_message(error.message())
}

pub(crate) fn to_abi_error(error: Error) -> AbiError {
    AbiError::from_hresult(HRESULT(error.code().raw()))
}

pub(crate) fn guid_to_abi(value: &Guid) -> GUID {
    GUID::from_values(value.data1, value.data2, value.data3, value.data4)
}

pub(crate) fn class_context_to_abi(value: ClassContext) -> CLSCTX {
    CLSCTX(value.bits())
}

pub(crate) fn object_from_interface<T: Interface>(value: &T) -> ComObject {
    // SAFETY: A windows-rs interface is non-null and cloning it transfers one
    // owned COM reference to `ComObject`.
    unsafe { ComObject::from_raw_owned(value.clone().into_raw()) }
        .expect("windows-rs produced a null COM interface")
}

pub(crate) fn interface_from_object<T: Interface>(value: &ComObject) -> Result<T> {
    let iid = T::IID;
    let iid = Guid::new(iid.data1, iid.data2, iid.data3, iid.data4);
    let requested = value.query_interface(&iid)?;
    // SAFETY: QueryInterface succeeded for `T::IID` and transferred an owned
    // reference through `ComObject::into_raw`.
    Ok(unsafe { T::from_raw(requested.into_raw()) })
}

pub(crate) fn timestamp_from_abi(value: FILETIME) -> Timestamp {
    Timestamp::from_ticks((u64::from(value.dwHighDateTime) << 32) | u64::from(value.dwLowDateTime))
}

pub(crate) fn timestamp_to_abi(value: Timestamp) -> FILETIME {
    let ticks = value.ticks();
    FILETIME {
        dwLowDateTime: ticks as u32,
        dwHighDateTime: (ticks >> 32) as u32,
    }
}

pub(crate) fn error_code_from_abi(value: HRESULT) -> ErrorCode {
    ErrorCode::from_raw(value.0)
}

pub(crate) fn error_code_to_abi(value: ErrorCode) -> HRESULT {
    HRESULT(value.raw())
}

fn scalar_variant(vt: VARENUM, value: VARIANT_0_0_0) -> VARIANT {
    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: value,
            }),
        },
    }
}

pub(crate) fn value_to_abi(value: &Value) -> Result<VARIANT> {
    match value {
        Value::Empty => Ok(VARIANT::default()),
        Value::Null => Ok(scalar_variant(VT_NULL, VARIANT_0_0_0::default())),
        Value::Bool(value) => Ok((*value).into()),
        Value::I8(value) => Ok((*value).into()),
        Value::U8(value) => Ok((*value).into()),
        Value::I16(value) => Ok((*value).into()),
        Value::U16(value) => Ok((*value).into()),
        Value::I32(value) => Ok((*value).into()),
        Value::U32(value) => Ok((*value).into()),
        Value::I64(value) => Ok((*value).into()),
        Value::U64(value) => Ok((*value).into()),
        Value::F32(value) => Ok((*value).into()),
        Value::F64(value) => Ok((*value).into()),
        Value::Currency(value) => Ok(scalar_variant(
            VT_CY,
            VARIANT_0_0_0 {
                cyVal: CY { int64: *value },
            },
        )),
        Value::Date(value) => Ok(scalar_variant(VT_DATE, VARIANT_0_0_0 { date: *value })),
        Value::String(value) => Ok(VARIANT::from(value.as_str())),
        Value::Error(value) => Ok(scalar_variant(
            VT_ERROR,
            VARIANT_0_0_0 { scode: value.raw() },
        )),
        Value::Bytes(_) | Value::Array(_) => Err(unsupported_value(
            "byte and array values are not supported by the DA ABI converter",
        )),
        _ => Err(unsupported_value(
            "value variant is not supported by the DA ABI converter",
        )),
    }
}

pub(crate) fn value_from_abi(value: &VARIANT) -> Result<Value> {
    let inner = unsafe { &value.Anonymous.Anonymous };
    let data = &inner.Anonymous;
    match inner.vt {
        VT_EMPTY => Ok(Value::Empty),
        VT_NULL => Ok(Value::Null),
        VT_BOOL => Ok(Value::Bool(unsafe { data.boolVal } != VARIANT_BOOL(0))),
        VT_I1 => Ok(Value::I8(unsafe { data.cVal })),
        VT_UI1 => Ok(Value::U8(unsafe { data.bVal })),
        VT_I2 => Ok(Value::I16(unsafe { data.iVal })),
        VT_UI2 => Ok(Value::U16(unsafe { data.uiVal })),
        VT_I4 | VT_INT => Ok(Value::I32(unsafe { data.lVal })),
        VT_UI4 | VT_UINT => Ok(Value::U32(unsafe { data.ulVal })),
        VT_I8 => Ok(Value::I64(unsafe { data.llVal })),
        VT_UI8 => Ok(Value::U64(unsafe { data.ullVal })),
        VT_R4 => Ok(Value::F32(unsafe { data.fltVal })),
        VT_R8 => Ok(Value::F64(unsafe { data.dblVal })),
        VT_CY => Ok(Value::Currency(unsafe { data.cyVal.int64 })),
        VT_DATE => Ok(Value::Date(unsafe { data.date })),
        VT_ERROR => Ok(Value::Error(ErrorCode::from_raw(unsafe { data.scode }))),
        VT_BSTR => {
            let bstr = unsafe { &data.bstrVal };
            String::from_utf16(bstr)
                .map(Value::String)
                .map_err(|_| Error::invalid_argument("VARIANT contains invalid UTF-16"))
        }
        other => Err(unsupported_value(format!(
            "unsupported VARIANT type 0x{:04X}",
            other.0
        ))),
    }
}

fn unsupported_value(message: impl Into<std::borrow::Cow<'static, str>>) -> Error {
    Error::new(ErrorKind::Automation, ErrorCode::TYPE_MISMATCH, message)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn supported_values_round_trip_without_loss() {
        let values = [
            Value::Empty,
            Value::Null,
            Value::Bool(true),
            Value::I8(-8),
            Value::U8(8),
            Value::I16(-16),
            Value::U16(16),
            Value::I32(-32),
            Value::U32(32),
            Value::I64(-64),
            Value::U64(64),
            Value::F32(1.25),
            Value::F64(-2.5),
            Value::Currency(-12345),
            Value::Date(45123.5),
            Value::String("hello".into()),
            Value::Error(ErrorCode::from_raw(-42)),
        ];

        for value in values {
            let raw = value_to_abi(&value).expect("supported value should encode");
            assert_eq!(
                value_from_abi(&raw).expect("supported value should decode"),
                value
            );
        }
    }

    #[test]
    fn unsupported_values_are_reported() {
        assert!(value_to_abi(&Value::Bytes(vec![1, 2])).is_err());
        assert!(value_to_abi(&Value::Array(vec![Value::I32(1)])).is_err());
    }

    #[test]
    fn timestamp_preserves_filetime_bits() {
        let timestamp = Timestamp::from_ticks(0x0123_4567_89ab_cdef);
        let raw = timestamp_to_abi(timestamp);
        assert_eq!(timestamp_from_abi(raw), timestamp);
    }
}
