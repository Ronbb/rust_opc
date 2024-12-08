use std::str::FromStr;

use windows::Win32::System::Com::CoTaskMemFree;

pub struct Array<T: Sized> {
    pointer: *mut T,
    len: usize,
}

impl<T: Sized> Array<T> {
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
}

impl<T: Sized> Drop for Array<T> {
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
        Ok(Self::new(Some(s.encode_utf16().chain(Some(0)).collect())))
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
}
