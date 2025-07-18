use std::ptr;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};

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
    /// Creates a shallow copy of the pointer.
    ///
    /// # Safety
    ///
    /// The caller must ensure that only one instance is passed to functions
    /// that will free the memory, to avoid double-free errors.
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
