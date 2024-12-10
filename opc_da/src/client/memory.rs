use std::str::FromStr;

use windows::core::PWSTR;
use windows::Win32::System::Com::CoTaskMemFree;

#[derive(Debug, Clone, PartialEq)]
pub struct RemoteArray<T: Sized> {
    pointer: *mut T,
    len: u32,
}

impl<T: Sized> RemoteArray<T> {
    #[inline(always)]
    pub fn new(len: u32) -> Self {
        Self {
            pointer: std::ptr::null_mut(),
            len,
        }
    }

    #[inline(always)]
    pub fn from_raw(pointer: *mut T, len: u32) -> Self {
        Self { pointer, len }
    }

    #[inline(always)]
    pub fn empty() -> Self {
        Self {
            pointer: std::ptr::null_mut(),
            len: 0,
        }
    }

    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.pointer
    }

    /// Returns a slice to the underlying array.  
    ///
    /// # Safety  
    /// The caller must ensure that the pointer is valid for reads  
    /// and points to an array of `len` elements.  
    #[inline(always)]
    pub fn as_slice(&self) -> &[T] {
        if self.pointer.is_null() || self.len == 0 {
            return &[];
        }

        let len = usize::try_from(self.len).unwrap_or(0);

        unsafe { core::slice::from_raw_parts(self.pointer, len) }
    }

    #[inline(always)]
    pub fn len(&self) -> u32 {
        self.len
    }

    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    #[inline(always)]
    pub fn as_mut_len_ptr(&mut self) -> *mut u32 {
        &mut self.len
    }

    #[inline(always)]
    pub fn set_len(&mut self, len: u32) {
        self.len = len;
    }
}

impl<T: Sized> Default for RemoteArray<T> {
    #[inline(always)]
    fn default() -> Self {
        Self::empty()
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
            inner: std::ptr::null_mut(),
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
    #[inline(always)]
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
    type Error = windows::core::Error;

    #[inline(always)]
    fn try_from(value: RemotePointer<u16>) -> Result<Self, Self::Error> {
        if value.inner.is_null() {
            return Err(windows::Win32::Foundation::E_POINTER.into());
        }

        Ok(unsafe { PWSTR(value.inner).to_string() }?)
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
    inner: Option<Box<T>>,
}

impl<T: Sized> LocalPointer<T> {
    #[inline(always)]
    pub fn new(value: Option<T>) -> Self {
        Self {
            inner: value.map(|v| Box::new(v)),
        }
    }

    #[inline(always)]
    pub fn from_box(value: Box<T>) -> Self {
        Self { inner: Some(value) }
    }

    #[inline(always)]
    pub fn as_ptr(&self) -> *const T {
        match &self.inner {
            Some(value) => value.as_ref() as *const T,
            None => std::ptr::null_mut(),
        }
    }

    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut T {
        match &mut self.inner {
            Some(value) => value.as_mut() as *mut T,
            None => std::ptr::null_mut(),
        }
    }

    #[inline(always)]
    pub fn into_inner(self) -> Option<T> {
        self.inner.map(|v| *v)
    }

    #[inline(always)]
    pub fn inner(&self) -> Option<&T> {
        self.inner.as_ref().map(|v| v.as_ref())
    }
}

impl FromStr for LocalPointer<Vec<u16>> {
    type Err = windows::core::HRESULT;

    #[inline(always)]
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(Self::from(s))
    }
}

impl From<&str> for LocalPointer<Vec<u16>> {
    #[inline(always)]
    fn from(s: &str) -> Self {
        Self::new(Some(s.encode_utf16().chain(Some(0)).collect()))
    }
}

impl From<&String> for LocalPointer<Vec<u16>> {
    #[inline(always)]
    fn from(s: &String) -> Self {
        Self::new(Some(s.encode_utf16().chain(Some(0)).collect()))
    }
}

impl From<&[String]> for LocalPointer<Vec<Vec<u16>>> {
    #[inline(always)]
    fn from(values: &[String]) -> Self {
        Self::new(Some(
            values
                .iter()
                .map(|s| s.encode_utf16().chain(Some(0)).collect())
                .collect(),
        ))
    }
}

impl<T> LocalPointer<Vec<T>> {
    #[inline(always)]
    pub fn len(&self) -> usize {
        match &self.inner {
            Some(values) => values.len(),
            None => 0,
        }
    }

    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        match &self.inner {
            Some(values) => values.is_empty(),
            None => true,
        }
    }

    #[inline(always)]
    pub fn as_array_ptr(&self) -> *const T {
        match &self.inner {
            Some(values) => values.as_ptr(),
            None => std::ptr::null(),
        }
    }

    #[inline(always)]
    pub fn as_mut_array_ptr(&mut self) -> *mut T {
        match &mut self.inner {
            Some(values) => values.as_mut_ptr(),
            None => std::ptr::null_mut(),
        }
    }
}

impl LocalPointer<Vec<Vec<u16>>> {
    #[inline(always)]
    pub fn as_pwstr_array(&self) -> Vec<windows::core::PWSTR> {
        match &self.inner {
            Some(values) => values
                .iter()
                .map(|value| windows::core::PWSTR(value.as_ptr() as _))
                .collect(),
            None => vec![windows::core::PWSTR::null()],
        }
    }

    #[inline(always)]
    pub fn as_pcwstr_array(&self) -> Vec<windows::core::PCWSTR> {
        match &self.inner {
            Some(values) => values
                .iter()
                .map(|value| windows::core::PCWSTR::from_raw(value.as_ptr() as _))
                .collect(),
            None => vec![windows::core::PCWSTR::null()],
        }
    }
}

impl LocalPointer<Vec<u16>> {
    #[inline(always)]
    pub fn as_pwstr(&self) -> windows::core::PWSTR {
        match &self.inner {
            Some(value) => windows::core::PWSTR(value.as_ptr() as _),
            None => windows::core::PWSTR::null(),
        }
    }

    #[inline(always)]
    pub fn as_pcwstr(&self) -> windows::core::PCWSTR {
        match &self.inner {
            Some(value) => windows::core::PCWSTR::from_raw(value.as_ptr() as _),
            None => windows::core::PCWSTR::null(),
        }
    }
}
