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
        // Keep the pointer intact for the callee to use
        // Note: We don't clear self.ptr because the callee needs it
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
    /// Creates a shallow copy of the pointer.
    ///
    /// # Safety
    ///
    /// The caller must ensure that only one instance is passed to functions
    /// that will free the memory, to avoid double-free errors.
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

    /// Creates a new `CalleeAllocatedPtr` from a value, allocating memory
    ///
    /// This allocates memory using `CoTaskMemAlloc` and copies the value into it.
    pub fn from_value(value: &T) -> Result<Self, windows::core::Error> {
        let size = std::mem::size_of::<T>();
        let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        unsafe {
            std::ptr::copy_nonoverlapping(value, ptr.cast(), 1);
        }
        Ok(unsafe { Self::new(ptr.cast()) })
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

/// Writes a value to a caller-allocated pointer using COM memory management
///
/// This macro simplifies writing values to caller-allocated pointers by wrapping
/// the raw pointer in a `CallerAllocatedPtr` and safely writing the value.
///
/// # Arguments
///
/// * `$ptr` - A raw pointer (`*mut T`) that points to caller-allocated memory
/// * `$value` - The value to write to the pointer
///
/// # Returns
///
/// Returns `Result<(), windows::core::Error>`:
/// * `Ok(())` - Value was successfully written
/// * `Err(E_INVALIDARG)` - The pointer is null or invalid
///
/// # Safety
///
/// The caller must ensure that:
/// * `$ptr` is a valid pointer to caller-allocated memory
/// * The memory pointed to by `$ptr` is properly initialized
/// * The callee (COM function) will be responsible for freeing the memory
///
/// # Example
///
/// ```rust
/// use opc_classic_utils::write_caller_allocated_ptr;
///
/// let mut count: u32 = 0;
/// let count_ptr = &mut count as *mut u32;
///
/// // Write a value to the caller-allocated pointer
/// write_caller_allocated_ptr!(count_ptr, 42u32)?;
/// assert_eq!(count, 42);
/// # Ok::<(), windows::core::Error>(())
/// ```
///
/// # Typical Use Cases
///
/// * Writing input parameters to COM function calls
/// * Setting up caller-allocated output parameters
/// * OPC Classic API parameter passing
#[macro_export]
macro_rules! write_caller_allocated_ptr {
    ($ptr:expr, $value:expr) => {
        unsafe {
            let mut ptr = opc_classic_utils::CallerAllocatedPtr::from_raw($ptr);
            let value = $value;
            ptr.as_mut()
                .ok_or(windows::Win32::Foundation::E_INVALIDARG)
                .map(|ptr| *ptr = value)
        }
    };
}
