use std::ptr;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};

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
        Self { ptr: self.ptr }
    }
}
