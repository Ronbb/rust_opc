//! Memory management utilities for the OPC DA client.
//!
//! This module provides safe wrappers around COM memory allocations and arrays.
//!
//! It includes three main types:
//! - `RemoteArray<T>` for managing arrays allocated by COM.
//! - `RemotePointer<T>` for managing single values allocated by COM.
//! - `LocalPointer<T>` for managing local memory that needs to be passed to COM functions.

use std::str::FromStr;
use windows::core::PWSTR;
use windows::Win32::System::Com::CoTaskMemFree;

/// A safe wrapper around arrays allocated by COM.
///
/// This struct ensures proper cleanup of COM-allocated memory when dropped.
/// It provides safe access to the underlying array through slices.
#[derive(Debug, Clone, PartialEq)]
pub struct RemoteArray<T: Sized> {
    pointer: *mut T,
    len: u32,
}

impl<T: Sized> RemoteArray<T> {
    /// Creates a new `RemoteArray` with the specified length.
    /// The underlying pointer is initialized to null.
    #[inline(always)]
    pub fn new(len: u32) -> Self {
        Self {
            pointer: std::ptr::null_mut(),
            len,
        }
    }

    /// Creates a `RemoteArray` from a raw pointer and length.
    ///
    /// # Safety
    /// The caller must ensure that the pointer is valid and points to a COM-allocated array.
    #[inline(always)]
    pub fn from_raw(pointer: *mut T, len: u32) -> Self {
        Self { pointer, len }
    }

    /// Creates an empty `RemoteArray`.
    #[inline(always)]
    pub fn empty() -> Self {
        Self {
            pointer: std::ptr::null_mut(),
            len: 0,
        }
    }

    /// Returns a mutable pointer to the array pointer.
    ///
    /// This is useful when calling COM functions that output an array via a pointer to a pointer.
    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.pointer
    }

    /// Returns a slice to the underlying array.
    ///
    /// # Safety
    /// The caller must ensure that the `pointer` is valid for reads and points to an array of `len` elements.
    #[inline(always)]
    pub fn as_slice(&self) -> &[T] {
        if self.pointer.is_null() || self.len == 0 {
            return &[];
        }

        let len = usize::try_from(self.len).unwrap_or(0);

        // Pointer and length are guaranteed to be valid
        unsafe { core::slice::from_raw_parts(self.pointer, len) }
    }

    /// Returns a mutable slice to the underlying array.
    ///
    /// # Safety
    /// The caller must ensure that the `pointer` is valid for reads and writes and points to an array of `len` elements.
    #[inline(always)]
    pub fn as_mut_slice(&mut self) -> &mut [T] {
        if self.pointer.is_null() || self.len == 0 {
            return &mut [];
        }

        let len = usize::try_from(self.len).unwrap_or(0);

        // Pointer and length are guaranteed to be valid
        unsafe { core::slice::from_raw_parts_mut(self.pointer, len) }
    }

    /// Returns the length of the array.
    #[inline(always)]
    pub fn len(&self) -> u32 {
        if self.pointer.is_null() {
            return 0;
        }

        self.len
    }

    /// Checks if the array is empty.
    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        self.len == 0 || self.pointer.is_null()
    }

    /// Returns a mutable pointer to the length.
    ///
    /// This is useful when calling COM functions that output the length via a pointer.
    #[inline(always)]
    pub fn as_mut_len_ptr(&mut self) -> *mut u32 {
        &mut self.len
    }

    /// Sets the length of the array.
    ///
    /// # Safety
    /// The caller must ensure that the new length is valid for the underlying array.
    #[inline(always)]
    pub(crate) unsafe fn set_len(&mut self, len: u32) {
        self.len = len;
    }
}

impl<T: Sized> Default for RemoteArray<T> {
    /// Creates an empty `RemoteArray` by default.
    #[inline(always)]
    fn default() -> Self {
        Self::empty()
    }
}

impl<T: Sized> Drop for RemoteArray<T> {
    /// Drops the `RemoteArray`, freeing the COM-allocated memory.
    #[inline(always)]
    fn drop(&mut self) {
        if !self.pointer.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.pointer as _));
            }
        }
    }
}

impl<T, E: Into<RemotePointer<T>> + Copy> From<RemoteArray<E>> for Vec<RemotePointer<T>> {
    /// Converts a `RemoteArray` to a vector of `RemotePointer`.
    #[inline(always)]
    fn from(array: RemoteArray<E>) -> Self {
        array
            .as_slice()
            .iter()
            .map(|value| (*value).into())
            .collect()
    }
}

/// A safe wrapper around a pointer allocated by COM.
///
/// This struct ensures proper cleanup of COM-allocated memory when dropped.
/// It provides methods to access the underlying pointer.
#[repr(transparent)]
pub struct RemotePointer<T: Sized> {
    inner: *mut T,
}

impl<T: Sized> RemotePointer<T> {
    /// Creates a new `RemotePointer` initialized to null.
    #[inline(always)]
    pub fn new() -> Self {
        Self {
            inner: std::ptr::null_mut(),
        }
    }

    /// Returns a mutable pointer to the inner pointer.
    ///
    /// Useful for COM functions that output data via a pointer to a pointer.
    #[inline(always)]
    pub(crate) fn from_raw(pointer: *mut T) -> Self {
        Self { inner: pointer }
    }

    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.inner
    }

    /// Returns an `Option` referencing the inner value if it is not null.
    ///
    /// # Safety
    /// The caller must ensure that the inner pointer is valid for reads.
    #[inline(always)]
    pub fn as_ref(&self) -> Option<&T> {
        // Pointer is guaranteed to be valid
        unsafe { self.inner.as_ref() }
    }

    #[inline(always)]
    pub fn as_result(&self) -> windows::core::Result<&T> {
        // Pointer is guaranteed to be valid
        unsafe { self.inner.as_ref() }.ok_or_else(|| {
            windows::core::Error::new(windows::Win32::Foundation::E_POINTER, "Pointer is null")
        })
    }
}

impl<T: Sized> Default for RemotePointer<T> {
    /// Creates a new `RemotePointer` initialized to null by default.
    #[inline(always)]
    fn default() -> Self {
        Self::new()
    }
}

impl From<PWSTR> for RemotePointer<u16> {
    /// Converts a `PWSTR` to a `RemotePointer<u16>`.
    #[inline(always)]
    fn from(value: PWSTR) -> Self {
        Self {
            inner: value.as_ptr(),
        }
    }
}

impl TryFrom<RemotePointer<u16>> for String {
    type Error = windows::core::Error;

    /// Attempts to convert a `RemotePointer<u16>` to a `String`.
    ///
    /// # Errors
    /// Returns an error if the pointer is null or if the string conversion fails.
    #[inline(always)]
    fn try_from(value: RemotePointer<u16>) -> Result<Self, Self::Error> {
        if value.inner.is_null() {
            return Err(windows::Win32::Foundation::E_POINTER.into());
        }

        // Has checked for null pointer
        Ok(unsafe { PWSTR(value.inner).to_string() }?)
    }
}

impl RemotePointer<u16> {
    /// Returns a mutable pointer to a `PWSTR`.
    #[inline(always)]
    pub fn as_mut_pwstr_ptr(&mut self) -> *mut PWSTR {
        &mut self.inner as *mut *mut u16 as *mut PWSTR
    }
}

impl<T: Sized> Drop for RemotePointer<T> {
    /// Drops the `RemotePointer`, freeing the COM-allocated memory.
    #[inline(always)]
    fn drop(&mut self) {
        if !self.inner.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.inner as _));
            }
        }
    }
}

/// A safe wrapper around locally allocated memory needing to be passed to COM functions.
///
/// This struct is useful for preparing data to be read by COM functions.
pub struct LocalPointer<T: Sized> {
    inner: Option<Box<T>>,
}

impl<T: Sized> LocalPointer<T> {
    /// Creates a new `LocalPointer` from an optional value.
    #[inline(always)]
    pub fn new(value: Option<T>) -> Self {
        Self {
            inner: value.map(|v| Box::new(v)),
        }
    }

    /// Creates a `LocalPointer` from a boxed value.
    #[inline(always)]
    pub fn from_box(value: Box<T>) -> Self {
        Self { inner: Some(value) }
    }

    /// Returns a constant pointer to the inner value.
    #[inline(always)]
    pub fn as_ptr(&self) -> *const T {
        match &self.inner {
            Some(value) => value.as_ref() as *const T,
            None => std::ptr::null_mut(),
        }
    }

    /// Returns a mutable pointer to the inner value.
    #[inline(always)]
    pub fn as_mut_ptr(&mut self) -> *mut T {
        match &mut self.inner {
            Some(value) => value.as_mut() as *mut T,
            None => std::ptr::null_mut(),
        }
    }

    /// Consumes the `LocalPointer`, returning the inner value if it exists.
    #[inline(always)]
    pub fn into_inner(self) -> Option<T> {
        self.inner.map(|v| *v)
    }

    /// Returns a reference to the inner value if it exists.
    #[inline(always)]
    pub fn inner(&self) -> Option<&T> {
        self.inner.as_ref().map(|v| v.as_ref())
    }
}

// Implementations for string handling

impl FromStr for LocalPointer<Vec<u16>> {
    type Err = windows::core::HRESULT;

    /// Converts a string slice to a `LocalPointer` containing a UTF-16 encoded null-terminated string.
    #[inline(always)]
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(Self::from(s))
    }
}

impl From<&str> for LocalPointer<Vec<u16>> {
    /// Converts a string slice to a `LocalPointer` containing a UTF-16 encoded null-terminated string.
    #[inline(always)]
    fn from(s: &str) -> Self {
        Self::new(Some(s.encode_utf16().chain(Some(0)).collect()))
    }
}

impl From<&String> for LocalPointer<Vec<u16>> {
    /// Converts a `String` reference to a `LocalPointer` containing a UTF-16 encoded null-terminated string.
    #[inline(always)]
    fn from(s: &String) -> Self {
        Self::new(Some(s.encode_utf16().chain(Some(0)).collect()))
    }
}

impl From<&[String]> for LocalPointer<Vec<Vec<u16>>> {
    /// Converts a slice of `String`s to a `LocalPointer` containing vectors of UTF-16 encoded null-terminated strings.
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
    /// Returns the length of the inner vector.
    #[inline(always)]
    pub fn len(&self) -> usize {
        match &self.inner {
            Some(values) => values.len(),
            None => 0,
        }
    }

    /// Checks if the inner vector is empty.
    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        match &self.inner {
            Some(values) => values.is_empty(),
            None => true,
        }
    }

    /// Returns a constant pointer to the inner array.
    #[inline(always)]
    pub fn as_array_ptr(&self) -> *const T {
        match &self.inner {
            Some(values) => values.as_ptr(),
            None => std::ptr::null(),
        }
    }

    /// Returns a mutable pointer to the inner array.
    #[inline(always)]
    pub fn as_mut_array_ptr(&mut self) -> *mut T {
        match &mut self.inner {
            Some(values) => values.as_mut_ptr(),
            None => std::ptr::null_mut(),
        }
    }
}

impl LocalPointer<Vec<Vec<u16>>> {
    /// Converts the inner vector of UTF-16 strings to a vector of `PWSTR`.
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

    /// Converts the inner vector of UTF-16 strings to a vector of `PCWSTR`.
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
    /// Converts the inner UTF-16 string to a `PWSTR`.
    #[inline(always)]
    pub fn as_pwstr(&self) -> windows::core::PWSTR {
        match &self.inner {
            Some(value) => windows::core::PWSTR(value.as_ptr() as _),
            None => windows::core::PWSTR::null(),
        }
    }

    /// Converts the inner UTF-16 string to a `PCWSTR`.
    #[inline(always)]
    pub fn as_pcwstr(&self) -> windows::core::PCWSTR {
        match &self.inner {
            Some(value) => windows::core::PCWSTR::from_raw(value.as_ptr() as _),
            None => windows::core::PCWSTR::null(),
        }
    }
}
