use windows::core::BSTR;
use windows::Win32::System::Variant::VARIANT;

use super::base::{AccessRight, Quality, Variant};

use opc_da_bindings;

impl Variant {
    // get type id
    pub fn get_data_type(&self) -> u16 {
        match self {
            Variant::Empty => windows::Win32::System::Variant::VT_EMPTY,
            Variant::Bool(_) => windows::Win32::System::Variant::VT_BOOL,
            Variant::String(_) => windows::Win32::System::Variant::VT_BSTR,
            Variant::I8(_) => windows::Win32::System::Variant::VT_I1,
            Variant::I16(_) => windows::Win32::System::Variant::VT_I2,
            Variant::I32(_) => windows::Win32::System::Variant::VT_I4,
            Variant::I64(_) => windows::Win32::System::Variant::VT_I8,
            Variant::F32(_) => windows::Win32::System::Variant::VT_R4,
            Variant::F64(_) => windows::Win32::System::Variant::VT_R8,
            Variant::U8(_) => windows::Win32::System::Variant::VT_UI1,
            Variant::U16(_) => windows::Win32::System::Variant::VT_UI2,
            Variant::U32(_) => windows::Win32::System::Variant::VT_UI4,
            Variant::U64(_) => windows::Win32::System::Variant::VT_UI8,
        }
        .0
    }
}

impl Quality {
    pub fn to_u16(&self) -> u16 {
        self.0
    }
}

impl AccessRight {
    pub fn to_u32(&self) -> u32 {
        let mut value = 0;
        if self.readable {
            value |= opc_da_bindings::OPC_READABLE;
        }
        if self.writable {
            value |= opc_da_bindings::OPC_WRITEABLE;
        }
        value
    }
}

impl From<Variant> for VARIANT {
    fn from(val: Variant) -> Self {
        match val {
            Variant::Empty => VARIANT::default(),
            Variant::Bool(value) => VARIANT::from(value),
            Variant::String(value) => VARIANT::from(BSTR::from(value)),
            Variant::I8(value) => VARIANT::from(value),
            Variant::I16(value) => VARIANT::from(value),
            Variant::I32(value) => VARIANT::from(value),
            Variant::I64(value) => VARIANT::from(value),
            Variant::F32(value) => VARIANT::from(value),
            Variant::F64(value) => VARIANT::from(value),
            Variant::U8(value) => VARIANT::from(value),
            Variant::U16(value) => VARIANT::from(value),
            Variant::U32(value) => VARIANT::from(value),
            Variant::U64(value) => VARIANT::from(value),
        }
    }
}

impl From<VARIANT> for Variant {
    fn from(value: VARIANT) -> Self {
        unsafe {
            let value = &value.Anonymous.Anonymous;
            match value.vt {
                windows::Win32::System::Variant::VT_EMPTY => Variant::Empty,
                windows::Win32::System::Variant::VT_BOOL => {
                    Variant::Bool(value.Anonymous.boolVal.as_bool())
                }
                windows::Win32::System::Variant::VT_BSTR => {
                    Variant::String(value.Anonymous.bstrVal.to_string())
                }
                windows::Win32::System::Variant::VT_I1 => Variant::I8(value.Anonymous.cVal),
                windows::Win32::System::Variant::VT_I2 => Variant::I16(value.Anonymous.iVal),
                windows::Win32::System::Variant::VT_I4 => Variant::I32(value.Anonymous.lVal),
                windows::Win32::System::Variant::VT_I8 => Variant::I64(value.Anonymous.llVal),
                windows::Win32::System::Variant::VT_R4 => Variant::F32(value.Anonymous.fltVal),
                windows::Win32::System::Variant::VT_R8 => Variant::F64(value.Anonymous.dblVal),
                windows::Win32::System::Variant::VT_UI1 => Variant::U8(value.Anonymous.bVal),
                windows::Win32::System::Variant::VT_UI2 => Variant::U16(value.Anonymous.uiVal),
                windows::Win32::System::Variant::VT_UI4 => Variant::U32(value.Anonymous.ulVal),
                windows::Win32::System::Variant::VT_UI8 => Variant::U64(value.Anonymous.ullVal),
                _ => Variant::Empty,
            }
        }
    }
}
