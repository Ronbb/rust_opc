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

macro_rules! impl_from_scalar {
    ($($type:ty => $variant:ident),+ $(,)?) => {
        $(
            impl From<$type> for Value {
                fn from(value: $type) -> Self {
                    Self::$variant(value)
                }
            }
        )+
    };
}

impl_from_scalar! {
    bool => Bool,
    i8 => I8,
    u8 => U8,
    i16 => I16,
    u16 => U16,
    i32 => I32,
    u32 => U32,
    i64 => I64,
    u64 => U64,
    f32 => F32,
    f64 => F64,
}

impl From<Vec<u8>> for Value {
    fn from(value: Vec<u8>) -> Self {
        Self::Bytes(value)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalar_and_byte_conversions_cover_the_owned_value_variants() {
        assert_eq!(Value::from(7_i16), Value::I16(7));
        assert_eq!(Value::from(7_u32), Value::U32(7));
        assert_eq!(Value::from(1.5_f32), Value::F32(1.5));
        assert_eq!(Value::from(vec![1_u8, 2]), Value::Bytes(vec![1, 2]));
    }
}
