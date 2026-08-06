use std::borrow::Cow;
use std::fmt;

/// An ABI-compatible operation status code.
///
/// Negative values represent failure. The numeric representation is retained
/// so vendor-specific OPC codes can cross the safe API without a Windows type.
#[repr(transparent)]
#[derive(Clone, Copy, Default, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct ErrorCode(i32);

impl ErrorCode {
    pub const OK: Self = Self(0);
    pub const PARTIAL_SUCCESS: Self = Self(1);
    pub const NOT_IMPLEMENTED: Self = Self(0x8000_4001_u32 as i32);
    pub const NO_INTERFACE: Self = Self(0x8000_4002_u32 as i32);
    pub const NULL_POINTER: Self = Self(0x8000_4003_u32 as i32);
    pub const OUT_OF_MEMORY: Self = Self(0x8007_000E_u32 as i32);
    pub const INVALID_ARGUMENT: Self = Self(0x8007_0057_u32 as i32);
    pub const UNEXPECTED: Self = Self(0x8000_FFFF_u32 as i32);
    pub const NO_AGGREGATION: Self = Self(0x8004_0110_u32 as i32);
    pub const CLASS_NOT_REGISTERED: Self = Self(0x8004_0154_u32 as i32);
    pub const TYPE_MISMATCH: Self = Self(0x8002_0005_u32 as i32);

    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    pub const fn raw(self) -> i32 {
        self.0
    }

    pub const fn is_success(self) -> bool {
        self.0 >= 0
    }

    pub const fn is_failure(self) -> bool {
        self.0 < 0
    }
}

impl fmt::Debug for ErrorCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "ErrorCode(0x{:08X})", self.0 as u32)
    }
}

impl fmt::Display for ErrorCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "0x{:08X}", self.0 as u32)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum ErrorKind {
    InvalidArgument,
    NullPointer,
    OutOfMemory,
    NotImplemented,
    NoInterface,
    Automation,
    Protocol,
    Panic,
    Other,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Error {
    code: ErrorCode,
    kind: ErrorKind,
    message: Cow<'static, str>,
}

impl Error {
    pub fn new(kind: ErrorKind, code: ErrorCode, message: impl Into<Cow<'static, str>>) -> Self {
        Self {
            code,
            kind,
            message: message.into(),
        }
    }

    pub fn from_code(code: ErrorCode) -> Self {
        let kind = match code {
            ErrorCode::INVALID_ARGUMENT => ErrorKind::InvalidArgument,
            ErrorCode::NULL_POINTER => ErrorKind::NullPointer,
            ErrorCode::OUT_OF_MEMORY => ErrorKind::OutOfMemory,
            ErrorCode::NOT_IMPLEMENTED => ErrorKind::NotImplemented,
            ErrorCode::NO_INTERFACE => ErrorKind::NoInterface,
            _ => ErrorKind::Other,
        };
        Self::new(kind, code, "OPC operation failed")
    }

    pub fn invalid_argument(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(
            ErrorKind::InvalidArgument,
            ErrorCode::INVALID_ARGUMENT,
            message,
        )
    }

    pub fn null_pointer(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(ErrorKind::NullPointer, ErrorCode::NULL_POINTER, message)
    }

    pub fn out_of_memory(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(ErrorKind::OutOfMemory, ErrorCode::OUT_OF_MEMORY, message)
    }

    pub fn unexpected(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(ErrorKind::Other, ErrorCode::UNEXPECTED, message)
    }

    pub fn not_implemented(message: impl Into<Cow<'static, str>>) -> Self {
        Self::new(
            ErrorKind::NotImplemented,
            ErrorCode::NOT_IMPLEMENTED,
            message,
        )
    }

    pub fn code(&self) -> ErrorCode {
        self.code
    }

    pub fn kind(&self) -> ErrorKind {
        self.kind
    }

    pub fn message(&self) -> &str {
        &self.message
    }

    /// Returns the status that is safe to emit when this value crosses a COM
    /// ABI boundary.
    ///
    /// `Error` normally carries a failing status. OPC batch methods also use
    /// `S_FALSE` (`PARTIAL_SUCCESS`) as a completion signal, so that one
    /// non-failing value is preserved. Any other success code stored in an
    /// `Error` is a contract violation and is normalized to `UNEXPECTED` to
    /// prevent a Rust `Err` from becoming an apparent COM success.
    pub fn boundary_status(&self) -> ErrorCode {
        match self.code {
            ErrorCode::PARTIAL_SUCCESS => ErrorCode::PARTIAL_SUCCESS,
            code if code.is_failure() => code,
            _ => ErrorCode::UNEXPECTED,
        }
    }

    pub fn with_message(mut self, message: impl Into<Cow<'static, str>>) -> Self {
        self.message = message.into();
        self
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{} ({})", self.message, self.code)
    }
}

impl std::error::Error for Error {}

pub type Result<T> = core::result::Result<T, Error>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn failure_uses_signed_status_bit() {
        assert!(ErrorCode::INVALID_ARGUMENT.is_failure());
        assert!(ErrorCode::OK.is_success());
    }

    #[test]
    fn boundary_status_rejects_success_codes_stored_as_errors() {
        assert_eq!(
            Error::from_code(ErrorCode::OK).boundary_status(),
            ErrorCode::UNEXPECTED
        );
        assert_eq!(
            Error::from_code(ErrorCode::from_raw(7)).boundary_status(),
            ErrorCode::UNEXPECTED
        );
        assert_eq!(
            Error::from_code(ErrorCode::INVALID_ARGUMENT).boundary_status(),
            ErrorCode::INVALID_ARGUMENT
        );
        assert_eq!(
            Error::from_code(ErrorCode::PARTIAL_SUCCESS).boundary_status(),
            ErrorCode::PARTIAL_SUCCESS
        );
    }

    #[test]
    fn constructors_and_known_codes_preserve_error_semantics() {
        assert_eq!(
            Error::from_code(ErrorCode::NULL_POINTER).kind(),
            ErrorKind::NullPointer
        );
        assert_eq!(
            Error::from_code(ErrorCode::TYPE_MISMATCH).kind(),
            ErrorKind::Other
        );

        let error = Error::out_of_memory("allocation failed");
        assert_eq!(error.code(), ErrorCode::OUT_OF_MEMORY);
        assert_eq!(error.kind(), ErrorKind::OutOfMemory);
        assert_eq!(error.message(), "allocation failed");
    }
}
