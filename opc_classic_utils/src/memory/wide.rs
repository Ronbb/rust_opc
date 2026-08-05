use std::ffi::{OsStr, OsString};
use std::fmt;
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::ptr::{self, NonNull};

use windows::Win32::Foundation::E_OUTOFMEMORY;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows_core::{PCWSTR, PWSTR};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct WideStringError;

impl fmt::Display for WideStringError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("wide string contains an interior NUL")
    }
}

impl std::error::Error for WideStringError {}

/// A Rust-owned, null-terminated UTF-16 string for borrowed COM input arguments.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WideCString(Vec<u16>);

impl WideCString {
    pub fn new(value: impl AsRef<OsStr>) -> Result<Self, WideStringError> {
        let mut wide: Vec<u16> = value.as_ref().encode_wide().collect();
        if wide.contains(&0) {
            return Err(WideStringError);
        }
        wide.push(0);
        Ok(Self(wide))
    }

    pub fn as_pcwstr(&self) -> PCWSTR {
        PCWSTR(self.0.as_ptr())
    }

    pub fn as_slice_with_nul(&self) -> &[u16] {
        &self.0
    }
}

impl TryFrom<&str> for WideCString {
    type Error = WideStringError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Owns a null-terminated `PWSTR` allocated with `CoTaskMemAlloc`.
pub struct OwnedPwstr(Option<NonNull<u16>>);

impl OwnedPwstr {
    /// Allocates a COM task-owned copy suitable for an ABI output parameter.
    pub fn new(value: impl AsRef<OsStr>) -> windows_core::Result<Self> {
        let wide = WideCString::new(value).map_err(|_| {
            windows_core::Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG)
        })?;
        let bytes = wide
            .0
            .len()
            .checked_mul(std::mem::size_of::<u16>())
            .ok_or_else(|| windows_core::Error::from_hresult(E_OUTOFMEMORY))?;
        let allocation = unsafe { CoTaskMemAlloc(bytes) }.cast::<u16>();
        let allocation = NonNull::new(allocation)
            .ok_or_else(|| windows_core::Error::from_hresult(E_OUTOFMEMORY))?;
        // SAFETY: Both allocations contain `wide.0.len()` consecutive `u16`s.
        unsafe {
            ptr::copy_nonoverlapping(wide.0.as_ptr(), allocation.as_ptr(), wide.0.len());
        }
        Ok(Self(Some(allocation)))
    }

    /// Adopts a COM task-allocated string.
    ///
    /// # Safety
    ///
    /// `ptr` must be null or a valid null-terminated UTF-16 string allocated with
    /// `CoTaskMemAlloc` and exclusively transferred to this value.
    pub unsafe fn from_raw(ptr: *mut u16) -> Self {
        Self(NonNull::new(ptr))
    }

    pub fn is_null(&self) -> bool {
        self.0.is_none()
    }

    pub fn as_pwstr(&self) -> PWSTR {
        PWSTR(self.0.map_or(ptr::null_mut(), |ptr| ptr.as_ptr()))
    }

    pub fn to_os_string(&self) -> OsString {
        let Some(ptr) = self.0 else {
            return OsString::new();
        };

        let mut len = 0usize;
        // SAFETY: `from_raw` requires a valid null-terminated string.
        unsafe {
            while *ptr.as_ptr().add(len) != 0 {
                len += 1;
            }
            OsString::from_wide(std::slice::from_raw_parts(ptr.as_ptr(), len))
        }
    }

    pub fn to_string_lossy(&self) -> String {
        self.to_os_string().to_string_lossy().into_owned()
    }

    pub fn into_raw(mut self) -> PWSTR {
        let ptr = self.0.take().map_or(ptr::null_mut(), |ptr| ptr.as_ptr());
        PWSTR(ptr)
    }
}

impl fmt::Debug for OwnedPwstr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_tuple("OwnedPwstr")
            .field(&self.to_string_lossy())
            .finish()
    }
}

impl Drop for OwnedPwstr {
    fn drop(&mut self) {
        if let Some(ptr) = self.0.take() {
            // SAFETY: Ownership is established by `from_raw`.
            unsafe { CoTaskMemFree(Some(ptr.as_ptr().cast())) };
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wide_c_string_rejects_interior_nul() {
        assert!(WideCString::try_from("a\0b").is_err());
    }

    #[test]
    fn wide_c_string_is_terminated() {
        let value = WideCString::try_from("OPC").unwrap();
        assert_eq!(
            value.as_slice_with_nul(),
            &[b'O' as u16, b'P' as u16, b'C' as u16, 0]
        );
    }
}
