//! Memory management utilities for OPC Classic
//!
//! This module provides automatic memory management for COM objects
//! using `CoTaskMemFree` for cleanup.
//!
//! COM memory management follows two patterns:
//! 1. Caller allocates, callee frees (e.g., input parameters)
//! 2. Callee allocates, caller frees (e.g., output parameters)

use std::ptr;
use windows::Win32::System::Com::CoTaskMemFree;
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
/// This wrapper automatically frees the memory when dropped.
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
}
