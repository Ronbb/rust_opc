use std::marker::PhantomData;
use std::ptr;

use windows::Win32::System::Com::CoTaskMemFree;

use super::{Cleanup, CoTaskMemArray, OwnedPwstr};

/// A null-initialized COM output pointer.
///
/// It releases an unclaimed outer allocation on drop, including when the COM
/// call returns an error. Use `into_array` or `into_pwstr` immediately after the
/// call to attach the correct element cleanup contract.
pub struct CoTaskMemOut<T> {
    ptr: *mut T,
    _type: PhantomData<T>,
}

impl<T> CoTaskMemOut<T> {
    pub const fn new() -> Self {
        Self {
            ptr: ptr::null_mut(),
            _type: PhantomData,
        }
    }

    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.ptr
    }

    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Attaches a length and element cleanup policy to the returned allocation.
    ///
    /// # Safety
    ///
    /// A non-null pointer must refer to `len` initialized values allocated with
    /// `CoTaskMemAlloc`, and `cleanup` must match those values.
    pub unsafe fn into_array<C: Cleanup<T>>(
        mut self,
        len: usize,
        cleanup: C,
    ) -> windows_core::Result<CoTaskMemArray<T, C>> {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        unsafe { CoTaskMemArray::from_raw_parts(ptr, len, cleanup) }
    }
}

impl CoTaskMemOut<u16> {
    /// Treats the returned allocation as a null-terminated `PWSTR`.
    ///
    /// # Safety
    ///
    /// The pointer must be null or point to a null-terminated UTF-16 string in a
    /// `CoTaskMemAlloc` allocation.
    pub unsafe fn into_pwstr(mut self) -> OwnedPwstr {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        unsafe { OwnedPwstr::from_raw(ptr) }
    }
}

impl<T> Default for CoTaskMemOut<T> {
    fn default() -> Self {
        Self::new()
    }
}

impl<T> Drop for CoTaskMemOut<T> {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            // SAFETY: `CoTaskMemOut` is only for output allocations governed by
            // the COM task allocator. Nested resources require `into_array`.
            unsafe { CoTaskMemFree(Some(self.ptr.cast())) };
            self.ptr = ptr::null_mut();
        }
    }
}
