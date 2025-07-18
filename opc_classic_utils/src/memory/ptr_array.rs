use std::ptr;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};

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
