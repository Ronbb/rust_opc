use crate::ErrorCode;

/// An Automation value type code without exposing `VARENUM`.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq)]
pub struct ValueType(u16);

impl ValueType {
    pub const EMPTY: Self = Self(0);
    pub const NULL: Self = Self(1);
    pub const I16: Self = Self(2);
    pub const I32: Self = Self(3);
    pub const F32: Self = Self(4);
    pub const F64: Self = Self(5);
    pub const CURRENCY: Self = Self(6);
    pub const DATE: Self = Self(7);
    pub const STRING: Self = Self(8);
    pub const ERROR: Self = Self(10);
    pub const BOOL: Self = Self(11);
    pub const I8: Self = Self(16);
    pub const U8: Self = Self(17);
    pub const U16: Self = Self(18);
    pub const U32: Self = Self(19);
    pub const I64: Self = Self(20);
    pub const U64: Self = Self(21);

    pub const fn from_raw(value: u16) -> Self {
        Self(value)
    }

    pub const fn raw(self) -> u16 {
        self.0
    }
}

/// A fully owned value used by safe OPC client and server APIs.
#[derive(Clone, Debug, Default, PartialEq)]
#[non_exhaustive]
pub enum Value {
    #[default]
    Empty,
    Null,
    Bool(bool),
    I8(i8),
    U8(u8),
    I16(i16),
    U16(u16),
    I32(i32),
    U32(u32),
    I64(i64),
    U64(u64),
    F32(f32),
    F64(f64),
    Currency(i64),
    Date(f64),
    String(String),
    Error(ErrorCode),
    Bytes(Vec<u8>),
    Array(Vec<Value>),
}

impl From<i32> for Value {
    fn from(value: i32) -> Self {
        Self::I32(value)
    }
}

impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Self::F64(value)
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::String(value.to_owned())
    }
}
