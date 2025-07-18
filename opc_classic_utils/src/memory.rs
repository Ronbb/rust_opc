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

/// A smart pointer for COM memory arrays that the **caller allocates and callee frees**
///
/// This is used for input array parameters where the caller allocates memory
/// and the callee (COM function) is responsible for freeing it.
/// This wrapper does NOT free the memory when dropped.
#[derive(Debug)]
pub struct CallerAllocatedArray<T> {
    ptr: *mut T,
    len: usize,
}

impl<T> CallerAllocatedArray<T> {
    /// Creates a new `CallerAllocatedArray` from a raw pointer and length
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer to an array of `len` elements
    /// allocated by the caller and that the callee will be responsible for freeing it.
    pub unsafe fn new(ptr: *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Creates a new `CallerAllocatedArray` from a raw pointer and length, taking ownership
    pub fn from_raw(ptr: *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Allocates memory for an array using `CoTaskMemAlloc` and creates a `CallerAllocatedArray`
    ///
    /// This allocates memory that will be freed by the callee (COM function).
    /// The caller is responsible for ensuring the callee will free this memory.
    pub fn allocate(len: usize) -> Result<Self, windows::core::Error> {
        if len == 0 {
            return Ok(Self {
                ptr: ptr::null_mut(),
                len: 0,
            });
        }

        let size = std::mem::size_of::<T>().checked_mul(len).ok_or_else(|| {
            windows::core::Error::new(
                windows::core::HRESULT::from_win32(0x80070057), // E_INVALIDARG
                "Array size overflow",
            )
        })?;

        let ptr = unsafe { CoTaskMemAlloc(size) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        Ok(unsafe { Self::new(ptr.cast(), len) })
    }

    /// Allocates memory and initializes it with a copy of the given slice
    ///
    /// This creates a copy of the slice in COM-allocated memory.
    pub fn from_slice(slice: &[T]) -> Result<Self, windows::core::Error>
    where
        T: Copy,
    {
        if slice.is_empty() {
            return Ok(Self {
                ptr: ptr::null_mut(),
                len: 0,
            });
        }

        let array = Self::allocate(slice.len())?;
        unsafe {
            std::ptr::copy_nonoverlapping(slice.as_ptr(), array.as_ptr(), slice.len());
        }
        Ok(array)
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CallerAllocatedArray` will not manage the memory.
    pub fn into_raw(mut self) -> (*mut T, usize) {
        let ptr = self.ptr;
        let len = self.len;
        self.ptr = ptr::null_mut();
        self.len = 0;
        (ptr, len)
    }

    /// Returns the length of the array
    pub fn len(&self) -> usize {
        self.len
    }

    /// Returns true if the array is empty
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Returns a slice of the array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_slice(&self) -> Option<&[T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts(self.ptr, self.len)) }
        }
    }

    /// Returns a mutable slice of the array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut_slice(&mut self) -> Option<&mut [T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts_mut(self.ptr, self.len)) }
        }
    }

    /// Gets an element at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get(&self, index: usize) -> Option<&T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(&*self.ptr.add(index)) }
        }
    }

    /// Gets a mutable element at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get_mut(&mut self, index: usize) -> Option<&mut T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(&mut *self.ptr.add(index)) }
        }
    }
}

impl<T> Drop for CallerAllocatedArray<T> {
    fn drop(&mut self) {
        // Do NOT free the memory - the callee is responsible for this
        // Just clear the pointer to prevent use-after-free
        self.ptr = ptr::null_mut();
        self.len = 0;
    }
}

impl<T> Default for CallerAllocatedArray<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: 0,
        }
    }
}

impl<T> Clone for CallerAllocatedArray<T> {
    fn clone(&self) -> Self {
        Self {
            ptr: self.ptr,
            len: self.len,
        }
    }
}

/// A smart pointer for COM memory arrays that the **callee allocates and caller frees**
///
/// This is used for output array parameters where the callee (COM function) allocates memory
/// and the caller is responsible for freeing it using `CoTaskMemFree`.
///
/// # Memory Management
/// - **Only frees the array container itself**
/// - **Does NOT free individual array elements**
/// - Use this when the callee returns an array of values (not pointers)
///
/// # Typical Use Cases
/// - OPC server returning an array of data values
/// - COM function returning an array of structures
/// - Any scenario where you receive a contiguous array of data
///
/// # Example
/// ```rust
/// use opc_classic_utils::memory::CalleeAllocatedArray;
/// use std::ptr;
///
/// // Server returns: [42.0, 84.0, 126.0] as *mut f64
/// let ptr = ptr::null_mut::<f64>(); // In real code, this would be from COM
/// let values = CalleeAllocatedArray::from_raw(ptr, 3);
/// // When values goes out of scope, only the array is freed
/// ```
#[derive(Debug)]
pub struct CalleeAllocatedArray<T> {
    ptr: *mut T,
    len: usize,
}

impl<T> CalleeAllocatedArray<T> {
    /// Creates a new `CalleeAllocatedArray` from a raw pointer and length
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer to an array of `len` elements
    /// allocated by the callee and that it will be freed using `CoTaskMemFree`.
    pub unsafe fn new(ptr: *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Creates a new `CalleeAllocatedArray` from a raw pointer and length, taking ownership
    ///
    /// This is safe when the pointer is null, as `CoTaskMemFree` handles null pointers.
    pub fn from_raw(ptr: *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CalleeAllocatedArray` will not free the memory.
    pub fn into_raw(mut self) -> (*mut T, usize) {
        let ptr = self.ptr;
        let len = self.len;
        self.ptr = ptr::null_mut();
        self.len = 0;
        (ptr, len)
    }

    /// Returns the length of the array
    pub fn len(&self) -> usize {
        self.len
    }

    /// Returns true if the array is empty
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Returns a slice of the array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_slice(&self) -> Option<&[T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts(self.ptr, self.len)) }
        }
    }

    /// Returns a mutable slice of the array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut_slice(&mut self) -> Option<&mut [T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts_mut(self.ptr, self.len)) }
        }
    }

    /// Gets an element at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get(&self, index: usize) -> Option<&T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(&*self.ptr.add(index)) }
        }
    }

    /// Gets a mutable element at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get_mut(&mut self, index: usize) -> Option<&mut T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(&mut *self.ptr.add(index)) }
        }
    }
}

impl<T> Drop for CalleeAllocatedArray<T> {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.ptr.cast()));
            }
            self.ptr = ptr::null_mut();
            self.len = 0;
        }
    }
}

impl<T> Default for CalleeAllocatedArray<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: 0,
        }
    }
}

impl<T> Clone for CalleeAllocatedArray<T> {
    fn clone(&self) -> Self {
        Self {
            ptr: self.ptr,
            len: self.len,
        }
    }
}

/// A smart pointer for COM memory pointer arrays that the **caller allocates and callee frees**
///
/// This is used for input pointer array parameters where the caller allocates memory
/// and the callee (COM function) is responsible for freeing it.
/// This wrapper does NOT free the memory when dropped.
#[derive(Debug)]
pub struct CallerAllocatedPtrArray<T> {
    ptr: *mut *mut T,
    len: usize,
}

impl<T> CallerAllocatedPtrArray<T> {
    /// Creates a new `CallerAllocatedPtrArray` from a raw pointer and length
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer to an array of `len` pointers
    /// allocated by the caller and that the callee will be responsible for freeing it.
    pub unsafe fn new(ptr: *mut *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Creates a new `CallerAllocatedPtrArray` from a raw pointer and length, taking ownership
    pub fn from_raw(ptr: *mut *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Allocates memory for a pointer array using `CoTaskMemAlloc` and creates a `CallerAllocatedPtrArray`
    ///
    /// This allocates memory that will be freed by the callee (COM function).
    /// The caller is responsible for ensuring the callee will free this memory.
    pub fn allocate(len: usize) -> Result<Self, windows::core::Error> {
        if len == 0 {
            return Ok(Self {
                ptr: ptr::null_mut(),
                len: 0,
            });
        }

        let size = std::mem::size_of::<*mut T>()
            .checked_mul(len)
            .ok_or_else(|| {
                windows::core::Error::new(
                    windows::core::HRESULT::from_win32(0x80070057), // E_INVALIDARG
                    "Pointer array size overflow",
                )
            })?;

        let ptr = unsafe { CoTaskMemAlloc(size) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        Ok(unsafe { Self::new(ptr.cast(), len) })
    }

    /// Allocates memory and initializes it with pointers from the given slice
    ///
    /// This creates a copy of the pointers in COM-allocated memory.
    pub fn from_ptr_slice(slice: &[*mut T]) -> Result<Self, windows::core::Error> {
        if slice.is_empty() {
            return Ok(Self {
                ptr: ptr::null_mut(),
                len: 0,
            });
        }

        let array = Self::allocate(slice.len())?;
        unsafe {
            std::ptr::copy_nonoverlapping(slice.as_ptr(), array.as_ptr(), slice.len());
        }
        Ok(array)
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CallerAllocatedPtrArray` will not manage the memory.
    pub fn into_raw(mut self) -> (*mut *mut T, usize) {
        let ptr = self.ptr;
        let len = self.len;
        self.ptr = ptr::null_mut();
        self.len = 0;
        (ptr, len)
    }

    /// Returns the length of the array
    pub fn len(&self) -> usize {
        self.len
    }

    /// Returns true if the array is empty
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Returns a slice of the pointer array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_slice(&self) -> Option<&[*mut T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts(self.ptr, self.len)) }
        }
    }

    /// Returns a mutable slice of the pointer array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut_slice(&mut self) -> Option<&mut [*mut T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts_mut(self.ptr, self.len)) }
        }
    }

    /// Gets a pointer at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get(&self, index: usize) -> Option<*mut T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(*self.ptr.add(index)) }
        }
    }

    /// Sets a pointer at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn set(&mut self, index: usize, value: *mut T) -> bool {
        if index >= self.len || self.ptr.is_null() {
            false
        } else {
            unsafe {
                *self.ptr.add(index) = value;
            }
            true
        }
    }
}

impl<T> Drop for CallerAllocatedPtrArray<T> {
    fn drop(&mut self) {
        // Do NOT free the memory - the callee is responsible for this
        // Just clear the pointer to prevent use-after-free
        self.ptr = ptr::null_mut();
        self.len = 0;
    }
}

impl<T> Default for CallerAllocatedPtrArray<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: 0,
        }
    }
}

impl<T> Clone for CallerAllocatedPtrArray<T> {
    fn clone(&self) -> Self {
        Self {
            ptr: self.ptr,
            len: self.len,
        }
    }
}

/// A smart pointer for COM memory pointer arrays that the **callee allocates and caller frees**
///
/// This is used for output pointer array parameters where the callee (COM function) allocates memory
/// and the caller is responsible for freeing it using `CoTaskMemFree`.
///
/// # Memory Management
/// - **Frees the array container itself**
/// - **ALSO frees each individual pointer in the array**
/// - Use this when the callee returns an array of pointers to allocated memory
///
/// # Typical Use Cases
/// - OPC server returning an array of string pointers
/// - COM function returning an array of object pointers
/// - Any scenario where you receive an array of pointers to allocated memory
///
/// # Example
/// ```rust
/// use opc_classic_utils::memory::CalleeAllocatedPtrArray;
/// use std::ptr;
///
/// // Server returns: [ptr1, ptr2, ptr3] where each ptr points to allocated memory
/// let ptr = ptr::null_mut::<*mut u16>(); // In real code, this would be from COM
/// let string_ptrs = CalleeAllocatedPtrArray::from_raw(ptr, 3);
/// // When string_ptrs goes out of scope:
/// // 1. Each pointer in the array is freed
/// // 2. The array itself is freed
/// ```
///
/// # ⚠️ Important Difference from CalleeAllocatedArray
/// - `CalleeAllocatedArray<T>`: Only frees the array container
/// - `CalleeAllocatedPtrArray<T>`: Frees both the array AND each pointer in it
#[derive(Debug)]
pub struct CalleeAllocatedPtrArray<T> {
    ptr: *mut *mut T,
    len: usize,
}

impl<T> CalleeAllocatedPtrArray<T> {
    /// Creates a new `CalleeAllocatedPtrArray` from a raw pointer and length
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid pointer to an array of `len` pointers
    /// allocated by the callee and that it will be freed using `CoTaskMemFree`.
    pub unsafe fn new(ptr: *mut *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Creates a new `CalleeAllocatedPtrArray` from a raw pointer and length, taking ownership
    ///
    /// This is safe when the pointer is null, as `CoTaskMemFree` handles null pointers.
    pub fn from_raw(ptr: *mut *mut T, len: usize) -> Self {
        Self { ptr, len }
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut *mut T {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    ///
    /// After calling this method, the `CalleeAllocatedPtrArray` will not free the memory.
    pub fn into_raw(mut self) -> (*mut *mut T, usize) {
        let ptr = self.ptr;
        let len = self.len;
        self.ptr = ptr::null_mut();
        self.len = 0;
        (ptr, len)
    }

    /// Returns the length of the array
    pub fn len(&self) -> usize {
        self.len
    }

    /// Returns true if the array is empty
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Returns a slice of the pointer array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_slice(&self) -> Option<&[*mut T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts(self.ptr, self.len)) }
        }
    }

    /// Returns a mutable slice of the pointer array if it's not null
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to initialized data.
    pub unsafe fn as_mut_slice(&mut self) -> Option<&mut [*mut T]> {
        if self.ptr.is_null() {
            None
        } else {
            unsafe { Some(std::slice::from_raw_parts_mut(self.ptr, self.len)) }
        }
    }

    /// Gets a pointer at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn get(&self, index: usize) -> Option<*mut T> {
        if index >= self.len || self.ptr.is_null() {
            None
        } else {
            unsafe { Some(*self.ptr.add(index)) }
        }
    }

    /// Sets a pointer at the given index
    ///
    /// # Safety
    ///
    /// The caller must ensure the index is within bounds and the pointer is valid.
    pub unsafe fn set(&mut self, index: usize, value: *mut T) -> bool {
        if index >= self.len || self.ptr.is_null() {
            false
        } else {
            unsafe {
                *self.ptr.add(index) = value;
            }
            true
        }
    }
}

impl<T> Drop for CalleeAllocatedPtrArray<T> {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe {
                // First, free each individual pointer in the array
                for i in 0..self.len {
                    let element_ptr = *self.ptr.add(i);
                    if !element_ptr.is_null() {
                        CoTaskMemFree(Some(element_ptr.cast()));
                    }
                }
                // Then free the array itself
                CoTaskMemFree(Some(self.ptr.cast()));
            }
            self.ptr = ptr::null_mut();
            self.len = 0;
        }
    }
}

impl<T> Default for CalleeAllocatedPtrArray<T> {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: 0,
        }
    }
}

impl<T> Clone for CalleeAllocatedPtrArray<T> {
    fn clone(&self) -> Self {
        Self {
            ptr: self.ptr,
            len: self.len,
        }
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

    #[test]
    fn test_caller_allocated_array_null() {
        let array = CallerAllocatedArray::<i32>::default();
        assert!(array.is_null());
        assert!(array.is_empty());
        assert_eq!(array.len(), 0);
    }

    #[test]
    fn test_callee_allocated_array_null() {
        let array = CalleeAllocatedArray::<i32>::default();
        assert!(array.is_null());
        assert!(array.is_empty());
        assert_eq!(array.len(), 0);
    }

    #[test]
    fn test_caller_allocated_array_allocate() {
        let array = CallerAllocatedArray::<i32>::allocate(5).unwrap();
        assert!(!array.is_null());
        assert!(!array.is_empty());
        assert_eq!(array.len(), 5);
    }

    #[test]
    fn test_caller_allocated_array_from_slice() {
        let data = vec![1, 2, 3, 4, 5];
        let array = CallerAllocatedArray::from_slice(&data).unwrap();
        assert!(!array.is_null());
        assert_eq!(array.len(), 5);

        // Verify the data was copied correctly
        unsafe {
            let slice = array.as_slice().unwrap();
            assert_eq!(slice, data.as_slice());
        }
    }

    #[test]
    fn test_caller_allocated_array_access() {
        let data = vec![10, 20, 30];
        let mut array = CallerAllocatedArray::from_slice(&data).unwrap();

        // Test get
        unsafe {
            assert_eq!(*array.get(0).unwrap(), 10);
            assert_eq!(*array.get(1).unwrap(), 20);
            assert_eq!(*array.get(2).unwrap(), 30);
            assert!(array.get(3).is_none()); // Out of bounds
        }

        // Test get_mut
        unsafe {
            *array.get_mut(1).unwrap() = 25;
            assert_eq!(*array.get(1).unwrap(), 25);
        }

        // Test as_slice and as_mut_slice
        unsafe {
            let slice = array.as_slice().unwrap();
            assert_eq!(slice, &[10, 25, 30]);

            let mut_slice = array.as_mut_slice().unwrap();
            mut_slice[2] = 35;
            assert_eq!(*array.get(2).unwrap(), 35);
        }
    }

    #[test]
    fn test_callee_allocated_array_no_free() {
        // This test verifies that CalleeAllocatedArray frees memory
        // In a real scenario, this would be memory allocated by the callee
        let _array: CalleeAllocatedArray<i32> =
            CalleeAllocatedArray::from_raw(std::ptr::null_mut(), 0);
        // When _array goes out of scope, it should call CoTaskMemFree
    }

    #[test]
    fn test_caller_allocated_ptr_array_null() {
        let array = CallerAllocatedPtrArray::<i32>::default();
        assert!(array.is_null());
        assert!(array.is_empty());
        assert_eq!(array.len(), 0);
    }

    #[test]
    fn test_callee_allocated_ptr_array_null() {
        let array = CalleeAllocatedPtrArray::<i32>::default();
        assert!(array.is_null());
        assert!(array.is_empty());
        assert_eq!(array.len(), 0);
    }

    #[test]
    fn test_caller_allocated_ptr_array_allocate() {
        let array = CallerAllocatedPtrArray::<i32>::allocate(3).unwrap();
        assert!(!array.is_null());
        assert!(!array.is_empty());
        assert_eq!(array.len(), 3);
    }

    #[test]
    fn test_caller_allocated_ptr_array_from_slice() {
        let ptrs = vec![
            std::ptr::null_mut::<i32>(),
            std::ptr::null_mut::<i32>(),
            std::ptr::null_mut::<i32>(),
        ];
        let array = CallerAllocatedPtrArray::from_ptr_slice(&ptrs).unwrap();
        assert!(!array.is_null());
        assert_eq!(array.len(), 3);

        // Verify the pointers were copied correctly
        unsafe {
            let slice = array.as_slice().unwrap();
            assert_eq!(slice, ptrs.as_slice());
        }
    }

    #[test]
    fn test_caller_allocated_ptr_array_access() {
        let mut array = CallerAllocatedPtrArray::<i32>::allocate(2).unwrap();

        // Test get and set
        unsafe {
            // Newly allocated memory contains uninitialized values, not necessarily null
            let _ptr0 = array.get(0).unwrap();
            let _ptr1 = array.get(1).unwrap();

            // Set to null and verify
            let test_ptr = std::ptr::null_mut::<i32>();
            assert!(array.set(0, test_ptr));
            assert_eq!(array.get(0).unwrap(), test_ptr);
            assert!(array.get(0).unwrap().is_null());

            // Test out of bounds
            assert!(!array.set(2, test_ptr)); // Out of bounds
        }
    }

    #[test]
    fn test_callee_allocated_ptr_array_no_free() {
        // This test verifies that CalleeAllocatedPtrArray frees memory
        // In a real scenario, this would be memory allocated by the callee
        let _array: CalleeAllocatedPtrArray<i32> =
            CalleeAllocatedPtrArray::from_raw(std::ptr::null_mut(), 0);
        // When _array goes out of scope, it should call CoTaskMemFree
    }

    #[test]
    fn test_array_transparent_repr() {
        // Test that transparent repr works correctly for arrays
        let ptr = std::ptr::null_mut::<i32>();
        let len = 5;

        // CallerAllocatedArray should have the same memory layout as (*mut i32, usize)
        let caller_array = CallerAllocatedArray::from_raw(ptr, len);
        assert_eq!(caller_array.as_ptr(), ptr);
        assert_eq!(caller_array.len(), len);

        // CalleeAllocatedArray should have the same memory layout as (*mut i32, usize)
        let callee_array = CalleeAllocatedArray::from_raw(ptr, len);
        assert_eq!(callee_array.as_ptr(), ptr);
        assert_eq!(callee_array.len(), len);

        // Pointer arrays should have the same memory layout as (*mut *mut i32, usize)
        let ptr_array = std::ptr::null_mut::<*mut i32>();
        let caller_ptr_array = CallerAllocatedPtrArray::from_raw(ptr_array, len);
        assert_eq!(caller_ptr_array.as_ptr(), ptr_array);
        assert_eq!(caller_ptr_array.len(), len);

        let callee_ptr_array = CalleeAllocatedPtrArray::from_raw(ptr_array, len);
        assert_eq!(callee_ptr_array.as_ptr(), ptr_array);
        assert_eq!(callee_ptr_array.len(), len);
    }
}
