use std::str::FromStr;

use windows::Win32::System::Com::CoTaskMemFree;
use windows_core::PWSTR;

pub struct RemoteArray<T: Sized> {
    pointer: *mut T,
    len: usize,
}

impl<T: Sized> RemoteArray<T> {
    #[inline(always)]
    pub fn new(len: usize) -> Self {
        Self {
            pointer: unsafe { core::mem::zeroed() },
            len,
        }
    }

    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.pointer
    }

    #[inline(always)]
    pub fn as_slice(&self) -> &[T] {
        if self.pointer.is_null() {
            return &[];
        }
        unsafe { core::slice::from_raw_parts(self.pointer, self.len) }
    }

    pub fn len(&self) -> usize {
        self.len
    }

    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    pub fn set_len(&mut self, len: usize) {
        self.len = len;
    }
}

impl<T: Sized> Drop for RemoteArray<T> {
    #[inline(always)]
    fn drop(&mut self) {
        if !self.pointer.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.pointer as _));
            }
        }
    }
}

#[repr(transparent)]
pub struct RemotePointer<T: Sized> {
    inner: *mut T,
}

impl<T: Sized> RemotePointer<T> {
    #[inline(always)]
    pub fn new() -> Self {
        Self {
            inner: unsafe { core::mem::zeroed() },
        }
    }

    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.inner
    }

    #[inline(always)]
    pub fn as_option(&self) -> Option<&T> {
        unsafe { self.inner.as_ref() }
    }
}

impl<T: Sized> Default for RemotePointer<T> {
    fn default() -> Self {
        Self::new()
    }
}

impl From<PWSTR> for RemotePointer<u16> {
    #[inline(always)]
    fn from(value: PWSTR) -> Self {
        Self {
            inner: value.as_ptr(),
        }
    }
}

impl TryFrom<RemotePointer<u16>> for String {
    type Error = windows_core::Error;

    fn try_from(value: RemotePointer<u16>) -> Result<Self, Self::Error> {
        if value.inner.is_null() {
            return Err(windows::Win32::Foundation::E_POINTER.into());
        }

        Ok(unsafe { PWSTR::from_raw(value.inner).to_string() }?)
    }
}

impl RemotePointer<u16> {
    #[inline(always)]
    pub fn as_mut_pwstr_ptr(&mut self) -> *mut PWSTR {
        &mut self.inner as *mut *mut u16 as *mut PWSTR
    }
}

impl<T: Sized> Drop for RemotePointer<T> {
    #[inline(always)]
    fn drop(&mut self) {
        if !self.inner.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.inner as _));
            }
        }
    }
}

pub struct LocalPointer<T: Sized> {
    inner: Option<T>,
}

impl<T: Sized> LocalPointer<T> {
    #[inline(always)]
    pub fn new(value: Option<T>) -> Self {
        Self { inner: value }
    }

    #[inline(always)]
    pub fn as_ptr(&self) -> *const T {
        match &self.inner {
            Some(value) => value as *const T,
            None => unsafe { core::mem::zeroed() },
        }
    }

    pub fn as_mut_ptr(&mut self) -> *mut T {
        match &mut self.inner {
            Some(value) => value as *mut T,
            None => unsafe { core::mem::zeroed() },
        }
    }

    #[inline(always)]
    pub fn into_inner(self) -> Option<T> {
        self.inner
    }
}

impl FromStr for LocalPointer<Vec<u16>> {
    type Err = windows_core::HRESULT;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(Self::from(s))
    }
}

impl From<&str> for LocalPointer<Vec<u16>> {
    fn from(s: &str) -> Self {
        Self::new(Some(s.encode_utf16().chain(Some(0)).collect()))
    }
}

impl From<&[String]> for LocalPointer<Vec<Vec<u16>>> {
    fn from(values: &[String]) -> Self {
        Self::new(Some(
            values
                .iter()
                .map(|s| s.encode_utf16().chain(Some(0)).collect())
                .collect(),
        ))
    }
}

impl LocalPointer<Vec<Vec<u16>>> {
    #[inline(always)]
    pub fn as_pwstr_array(&self) -> Vec<windows_core::PWSTR> {
        match &self.inner {
            Some(values) => values
                .iter()
                .map(|value| windows_core::PWSTR::from_raw(value.as_ptr() as _))
                .collect(),
            None => vec![windows_core::PWSTR::null()],
        }
    }

    #[inline(always)]
    pub fn as_pcwstr_array(&self) -> Vec<windows_core::PCWSTR> {
        match &self.inner {
            Some(values) => values
                .iter()
                .map(|value| windows_core::PCWSTR::from_raw(value.as_ptr() as _))
                .collect(),
            None => vec![windows_core::PCWSTR::null()],
        }
    }
}

impl LocalPointer<Vec<u16>> {
    #[inline(always)]
    pub fn as_pwstr(&self) -> windows_core::PWSTR {
        match &self.inner {
            Some(value) => windows_core::PWSTR::from_raw(value.as_ptr() as _),
            None => windows_core::PWSTR::null(),
        }
    }

    #[inline(always)]
    pub fn as_pcwstr(&self) -> windows_core::PCWSTR {
        match &self.inner {
            Some(value) => windows_core::PCWSTR::from_raw(value.as_ptr() as _),
            None => windows_core::PCWSTR::null(),
        }
    }
}
