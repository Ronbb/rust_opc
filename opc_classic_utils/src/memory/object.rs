use std::ops::{Deref, DerefMut};

use opc_classic_types::Result;

use super::{Cleanup, CoTaskMemArray};

/// Owns one initialized object allocated with the COM task allocator.
///
/// The cleanup policy releases resources nested inside `T`; the outer
/// allocation is released with `CoTaskMemFree` by the underlying owner.
pub struct CoTaskMemObject<T, C: Cleanup<T>> {
    inner: CoTaskMemArray<T, C>,
}

impl<T, C: Cleanup<T>> CoTaskMemObject<T, C> {
    /// Adopts one object returned by COM.
    ///
    /// # Safety
    ///
    /// `ptr` must point to one initialized `T` in a `CoTaskMemAlloc`
    /// allocation. `cleanup` must match the ownership contract of the object.
    pub unsafe fn from_raw(ptr: *mut T, cleanup: C) -> Result<Self> {
        let inner = unsafe { CoTaskMemArray::from_raw_parts(ptr, 1, cleanup) }?;
        Ok(Self { inner })
    }

    pub(crate) fn from_array(inner: CoTaskMemArray<T, C>) -> Self {
        debug_assert_eq!(inner.len(), 1);
        Self { inner }
    }

    pub fn as_ptr(&self) -> *const T {
        self.inner.as_ptr()
    }

    pub fn as_mut_ptr(&mut self) -> *mut T {
        self.inner.as_mut_ptr()
    }

    /// Transfers ownership of the allocation to the caller.
    pub fn into_raw(self) -> *mut T {
        let (ptr, len) = self.inner.into_raw_parts();
        debug_assert_eq!(len, 1);
        ptr
    }
}

impl<T, C: Cleanup<T>> AsRef<T> for CoTaskMemObject<T, C> {
    fn as_ref(&self) -> &T {
        &self.inner.as_slice()[0]
    }
}

impl<T, C: Cleanup<T>> AsMut<T> for CoTaskMemObject<T, C> {
    fn as_mut(&mut self) -> &mut T {
        &mut self.inner.as_mut_slice()[0]
    }
}

impl<T, C: Cleanup<T>> Deref for CoTaskMemObject<T, C> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        self.as_ref()
    }
}

impl<T, C: Cleanup<T>> DerefMut for CoTaskMemObject<T, C> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        self.as_mut()
    }
}
