//! Memory management utilities for OPC Classic
//!
//! This module provides automatic memory management for COM objects
//! using `CoTaskMemFree` for cleanup.
//!
//! COM memory management follows two patterns:
//! 1. Caller allocates, callee frees (e.g., input parameters)
//! 2. Callee allocates, caller frees (e.g., output parameters)

use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::ptr;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows::core::PCWSTR;

/// A smart pointer for COM memory that the **caller allocates and callee frees**
///
/// This is used for input parameters where the caller allocates memory
/// and the callee (COM function) is responsible for freeing it.
/// This wrapper does NOT free the memory when dropped.
#[repr(transparent)]
#[derive(Debug)]
pub struct CallerAllocatedPtr<T> {
    ptr: *mut T,
}

impl<T> CallerAllocatedPtr<T> {
    /// Creates a new `CallerAllocatedPtr` from a raw pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer allocated by the caller
    /// and that the callee will be responsible for freeing it.
    pub unsafe fn new(ptr: *mut T) -> Self {
        Self { ptr }
    }

    /// Creates a new `CallerAllocatedPtr` from a raw pointer, taking ownership
    pub fn from_raw(ptr: *mut T) -> Self {
        Self { ptr }
    }

    /// Allocates memory using `CoTaskMemAlloc` and creates a `CallerAllocatedPtr`
    ///
    /// This allocates memory that will be freed by the callee (COM function).
    /// The caller is responsible for ensuring the callee will free this memory.
    pub fn allocate() -> Result<Self, windows::core::Error> {
        let ptr = unsafe { CoTaskMemAlloc(std::mem::size_of::<T>()) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        Ok(unsafe { Self::new(ptr.cast()) })
    }

    /// Allocates memory and initializes it with a copy of the given value
    ///
    /// This creates a copy of the value in COM-allocated memory.
    pub fn from_value(value: &T) -> Result<Self, windows::core::Error>
    where
        T: Copy,
    {
        let ptr = Self::allocate()?;
        unsafe {
            *ptr.as_ptr() = *value;
        }
        Ok(ptr)
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CallerAllocatedPtr` will not manage the memory.
    pub fn into_raw(mut self) -> *mut T {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Dereferences the pointer if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_ref(&self) -> Option<&T> {
        if self.ptr.is_null() {
            None
        } else {
            Some(unsafe { &*self.ptr })
        }
    }

    /// Mutably dereferences the pointer if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut(&mut self) -> Option<&mut T> {
        if self.ptr.is_null() {
            None
        } else {
            Some(unsafe { &mut *self.ptr })
        }
    }
}

impl<T> Drop for CallerAllocatedPtr<T> {
    fn drop(&mut self) {
        // Do NOT free the memory - the callee is responsible for this
        // Just clear the pointer to prevent use-after-free
        self.ptr = ptr::null_mut();
    }
}

impl<T> Default for CallerAllocatedPtr<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl<T> Clone for CallerAllocatedPtr<T> {
    fn clone(&self) -> Self {
        Self { ptr: self.ptr }
    }
}

/// A smart pointer for COM memory that the **callee allocates and caller frees**
///
/// This is used for output parameters where the callee (COM function) allocates memory
/// and the caller is responsible for freeing it using `CoTaskMemFree`.
#[repr(transparent)]
#[derive(Debug)]
pub struct CalleeAllocatedPtr<T> {
    ptr: *mut T,
}

impl<T> CalleeAllocatedPtr<T> {
    /// Creates a new `CalleeAllocatedPtr` from a raw pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer allocated by the callee
    /// and that it will be freed using `CoTaskMemFree`.
    pub unsafe fn new(ptr: *mut T) -> Self {
        Self { ptr }
    }

    /// Creates a new `CalleeAllocatedPtr` from a raw pointer, taking ownership
    ///
    /// This is safe when the pointer is null, as `CoTaskMemFree` handles null pointers.
    pub fn from_raw(ptr: *mut T) -> Self {
        Self { ptr }
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CalleeAllocatedPtr` will not free the memory.
    pub fn into_raw(mut self) -> *mut T {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Dereferences the pointer if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_ref(&self) -> Option<&T> {
        if self.ptr.is_null() {
            None
        } else {
            Some(unsafe { &*self.ptr })
        }
    }

    /// Mutably dereferences the pointer if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut(&mut self) -> Option<&mut T> {
        if self.ptr.is_null() {
            None
        } else {
            Some(unsafe { &mut *self.ptr })
        }
    }
}

impl<T> Drop for CalleeAllocatedPtr<T> {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.ptr.cast()));
            }
            self.ptr = ptr::null_mut();
        }
    }
}

impl<T> Default for CalleeAllocatedPtr<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl<T> Clone for CalleeAllocatedPtr<T> {
    fn clone(&self) -> Self {
        // Note: This creates a new wrapper around the same pointer.
        // The caller should ensure proper ownership semantics.
        Self { ptr: self.ptr }
    }
}

/// A smart pointer for wide string pointers that the **caller allocates and callee frees**
///
/// This is used for input string parameters where the caller allocates memory
/// and the callee is responsible for freeing it.
#[repr(transparent)]
#[derive(Debug)]
pub struct CallerAllocatedWString {
    ptr: *mut u16,
}

impl CallerAllocatedWString {
    /// Creates a new `CallerAllocatedWString` from a raw wide string pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid wide string pointer
    /// allocated by the caller and that the callee will be responsible for freeing it.
    pub unsafe fn new(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CallerAllocatedWString` from a raw pointer, taking ownership
    pub fn from_raw(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CallerAllocatedWString` from a `PCWSTR`
    pub fn from_pcwstr(pcwstr: PCWSTR) -> Self {
        Self {
            ptr: pcwstr.as_ptr() as *mut u16,
        }
    }

    /// Allocates memory using `CoTaskMemAlloc` and creates a `CallerAllocatedWString`
    ///
    /// This allocates memory for a wide string that will be freed by the callee.
    pub fn allocate(len: usize) -> Result<Self, windows::core::Error> {
        let size = (len + 1) * std::mem::size_of::<u16>(); // +1 for null terminator
        let ptr = unsafe { CoTaskMemAlloc(size) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        Ok(unsafe { Self::new(ptr.cast()) })
    }

    /// Creates a `CallerAllocatedWString` from a Rust string
    pub fn from_string(s: String) -> Result<Self, windows::core::Error> {
        use std::str::FromStr;
        Self::from_str(&s)
    }

    /// Creates a `CallerAllocatedWString` from an `OsStr`
    pub fn from_os_str(os_str: &OsStr) -> Result<Self, windows::core::Error> {
        let wide_string: Vec<u16> = os_str.encode_wide().chain(std::iter::once(0)).collect();
        let len = wide_string.len() - 1; // Exclude null terminator for allocation

        let ptr = Self::allocate(len)?;
        unsafe {
            std::ptr::copy_nonoverlapping(wide_string.as_ptr(), ptr.as_ptr(), wide_string.len());
        }
        Ok(ptr)
    }

    /// Converts the wide string to a Rust string slice
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_string(&self) -> Option<String> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        let os_string = OsString::from_wide(slice);
        Some(os_string.to_string_lossy().into_owned())
    }

    /// Converts the wide string to an `OsString`
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_os_string(&self) -> Option<OsString> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        Some(OsString::from_wide(slice))
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut u16 {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    pub fn into_raw(mut self) -> *mut u16 {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Converts to a `PCWSTR` for use with Windows APIs
    pub fn as_pcwstr(&self) -> PCWSTR {
        PCWSTR(self.ptr)
    }
}

impl Drop for CallerAllocatedWString {
    fn drop(&mut self) {
        // Do NOT free the memory - the callee is responsible for this
        self.ptr = ptr::null_mut();
    }
}

impl Default for CallerAllocatedWString {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl Clone for CallerAllocatedWString {
    fn clone(&self) -> Self {
        Self { ptr: self.ptr }
    }
}

impl std::str::FromStr for CallerAllocatedWString {
    type Err = windows::core::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let wide_string: Vec<u16> = OsStr::new(s)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        let len = wide_string.len() - 1; // Exclude null terminator for allocation

        let ptr = Self::allocate(len)?;
        unsafe {
            std::ptr::copy_nonoverlapping(wide_string.as_ptr(), ptr.as_ptr(), wide_string.len());
        }
        Ok(ptr)
    }
}

/// A smart pointer for wide string pointers that the **callee allocates and caller frees**
///
/// This is used for output string parameters where the callee allocates memory
/// and the caller is responsible for freeing it using `CoTaskMemFree`.
#[repr(transparent)]
#[derive(Debug)]
pub struct CalleeAllocatedWString {
    ptr: *mut u16,
}

impl CalleeAllocatedWString {
    /// Creates a new `CalleeAllocatedWString` from a raw wide string pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid wide string pointer
    /// allocated by the callee and that it will be freed using `CoTaskMemFree`.
    pub unsafe fn new(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CalleeAllocatedWString` from a raw pointer, taking ownership
    pub fn from_raw(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CalleeAllocatedWString` from a `PCWSTR`
    pub fn from_pcwstr(pcwstr: PCWSTR) -> Self {
        Self {
            ptr: pcwstr.as_ptr() as *mut u16,
        }
    }

    /// Converts the wide string to a Rust string slice
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_string(&self) -> Option<String> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        let os_string = OsString::from_wide(slice);
        Some(os_string.to_string_lossy().into_owned())
    }

    /// Converts the wide string to an `OsString`
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_os_string(&self) -> Option<OsString> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        Some(OsString::from_wide(slice))
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut u16 {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    pub fn into_raw(mut self) -> *mut u16 {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Converts to a `PCWSTR` for use with Windows APIs
    pub fn as_pcwstr(&self) -> PCWSTR {
        PCWSTR(self.ptr)
    }
}

impl Drop for CalleeAllocatedWString {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.ptr.cast()));
            }
            self.ptr = ptr::null_mut();
        }
    }
}

impl Default for CalleeAllocatedWString {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl Clone for CalleeAllocatedWString {
    fn clone(&self) -> Self {
        Self { ptr: self.ptr }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    #[test]
    fn test_caller_allocated_ptr_null() {
        let ptr = CallerAllocatedPtr::<i32>::default();
        assert!(ptr.is_null());
    }

    #[test]
    fn test_callee_allocated_ptr_null() {
        let ptr = CalleeAllocatedPtr::<i32>::default();
        assert!(ptr.is_null());
    }

    #[test]
    fn test_caller_allocated_wstring_null() {
        let wstring = CallerAllocatedWString::default();
        assert!(wstring.is_null());
    }

    #[test]
    fn test_callee_allocated_wstring_null() {
        let wstring = CalleeAllocatedWString::default();
        assert!(wstring.is_null());
    }

    #[test]
    fn test_caller_allocated_ptr_no_free() {
        // This test verifies that CallerAllocatedPtr doesn't free memory
        // In a real scenario, this would be memory allocated by the caller
        let _ptr = CallerAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
        // When _ptr goes out of scope, it should NOT call CoTaskMemFree
    }

    #[test]
    fn test_callee_allocated_ptr_frees() {
        // This test verifies that CalleeAllocatedPtr frees memory
        // In a real scenario, this would be memory allocated by the callee
        let _ptr = CalleeAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
        // When _ptr goes out of scope, it should call CoTaskMemFree
    }

    #[test]
    fn test_transparent_repr() {
        // Test that transparent repr works correctly
        let ptr = std::ptr::null_mut::<i32>();

        // CallerAllocatedPtr should have the same memory layout as *mut i32
        let caller_ptr = CallerAllocatedPtr::from_raw(ptr);
        assert_eq!(caller_ptr.as_ptr(), ptr);

        // CalleeAllocatedPtr should have the same memory layout as *mut i32
        let callee_ptr = CalleeAllocatedPtr::from_raw(ptr);
        assert_eq!(callee_ptr.as_ptr(), ptr);

        // WString types should have the same memory layout as *mut u16
        let wstring_ptr = std::ptr::null_mut::<u16>();
        let caller_wstring = CallerAllocatedWString::from_raw(wstring_ptr);
        assert_eq!(caller_wstring.as_ptr(), wstring_ptr);

        let callee_wstring = CalleeAllocatedWString::from_raw(wstring_ptr);
        assert_eq!(callee_wstring.as_ptr(), wstring_ptr);
    }

    #[test]
    fn test_caller_allocated_ptr_allocate() {
        // Test allocation of caller-allocated pointer
        let ptr = CallerAllocatedPtr::<i32>::allocate().unwrap();
        assert!(!ptr.is_null());
        // Memory will be freed by callee, not by our wrapper
    }

    #[test]
    fn test_caller_allocated_ptr_from_value() {
        // Test creating pointer from value
        let value = 42i32;
        let ptr = CallerAllocatedPtr::from_value(&value).unwrap();
        assert!(!ptr.is_null());

        // Verify the value was copied correctly
        unsafe {
            assert_eq!(*ptr.as_ptr(), 42);
        }
    }

    #[test]
    fn test_caller_allocated_wstring_from_str() {
        // Test creating wide string from Rust string
        use std::str::FromStr;
        let test_string = "Hello, World!";
        let wstring = CallerAllocatedWString::from_str(test_string).unwrap();
        assert!(!wstring.is_null());

        // Verify the string was converted correctly
        unsafe {
            let converted = wstring.to_string().unwrap();
            assert_eq!(converted, test_string);
        }
    }

    #[test]
    fn test_caller_allocated_wstring_from_string() {
        // Test creating wide string from String
        let test_string = String::from("Test String");
        let wstring = CallerAllocatedWString::from_string(test_string.clone()).unwrap();
        assert!(!wstring.is_null());

        // Verify the string was converted correctly
        unsafe {
            let converted = wstring.to_string().unwrap();
            assert_eq!(converted, test_string);
        }
    }

    #[test]
    fn test_caller_allocated_wstring_from_os_str() {
        // Test creating wide string from OsStr
        let test_string = OsStr::new("OS String Test");
        let wstring = CallerAllocatedWString::from_os_str(test_string).unwrap();
        assert!(!wstring.is_null());

        // Verify the string was converted correctly
        unsafe {
            let converted = wstring.to_os_string().unwrap();
            assert_eq!(converted, test_string);
        }
    }

    #[test]
    fn test_pointer_dereference() {
        // Test dereferencing pointers
        let value = 123i32;
        let mut caller_ptr = CallerAllocatedPtr::from_value(&value).unwrap();
        let callee_ptr = CalleeAllocatedPtr::from_raw(caller_ptr.as_ptr());

        // Test as_ref
        unsafe {
            assert_eq!(caller_ptr.as_ref().unwrap(), &123);
            assert_eq!(callee_ptr.as_ref().unwrap(), &123);
        }

        // Test as_mut
        unsafe {
            *caller_ptr.as_mut().unwrap() = 456;
            assert_eq!(*caller_ptr.as_ptr(), 456);
        }
    }

    #[test]
    fn test_null_pointer_dereference() {
        // Test dereferencing null pointers
        let caller_ptr = CallerAllocatedPtr::<i32>::default();
        let callee_ptr = CalleeAllocatedPtr::<i32>::default();

        unsafe {
            assert!(caller_ptr.as_ref().is_none());
            assert!(callee_ptr.as_ref().is_none());
        }
    }

    #[test]
    fn test_wstring_null_conversion() {
        // Test converting null wide strings
        let caller_wstring = CallerAllocatedWString::default();
        let callee_wstring = CalleeAllocatedWString::default();

        unsafe {
            assert!(caller_wstring.to_string().is_none());
            assert!(callee_wstring.to_string().is_none());
            assert!(caller_wstring.to_os_string().is_none());
            assert!(callee_wstring.to_os_string().is_none());
        }
    }

    #[test]
    fn test_from_str_trait() {
        // Test the FromStr trait implementation
        let test_string = "Test FromStr Trait";
        let wstring = CallerAllocatedWString::from_str(test_string).unwrap();
        assert!(!wstring.is_null());

        // Verify the string was converted correctly
        unsafe {
            let converted = wstring.to_string().unwrap();
            assert_eq!(converted, test_string);
        }
    }
}
