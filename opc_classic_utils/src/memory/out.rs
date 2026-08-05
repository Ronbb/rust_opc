use std::marker::PhantomData;
use std::ptr;

use opc_classic_types::Result;
use windows::Win32::System::Com::CoTaskMemFree;

use super::{Cleanup, CoTaskMemArray, OwnedPwstr};

/// A null-initialized COM array output with its cleanup contract attached from
/// construction time.
///
/// Unlike [`CoTaskMemOut`], this guard knows how many initialized elements it
/// owns and how to release their nested resources. Dropping it before
/// [`Self::into_array`] therefore performs the same deep cleanup as the final
/// array owner.
pub struct CoTaskMemArrayOut<T, C: Cleanup<T>> {
    ptr: *mut T,
    len: usize,
    cleanup: Option<C>,
    _type: PhantomData<T>,
}

impl<T, C: Cleanup<T>> CoTaskMemArrayOut<T, C> {
    /// Creates an output guard for an array whose initialized length is known.
    pub const fn new(len: usize, cleanup: C) -> Self {
        Self {
            ptr: ptr::null_mut(),
            len,
            cleanup: Some(cleanup),
            _type: PhantomData,
        }
    }

    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.ptr
    }

    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Updates the initialized length reported by the foreign call.
    ///
    /// # Safety
    ///
    /// If the output pointer is non-null, it must address at least `len`
    /// initialized values governed by this guard's cleanup policy.
    pub unsafe fn set_len(&mut self, len: usize) {
        self.len = len;
    }

    /// Transfers the returned allocation into the final array owner.
    ///
    /// # Safety
    ///
    /// A non-null pointer must refer to `self.len` initialized values allocated
    /// with `CoTaskMemAlloc`, and the cleanup policy supplied to [`Self::new`]
    /// must match those values.
    pub unsafe fn into_array(mut self) -> Result<CoTaskMemArray<T, C>> {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        let cleanup = self
            .cleanup
            .take()
            .expect("array output cleanup policy is always present");
        unsafe { CoTaskMemArray::from_raw_parts(ptr, self.len, cleanup) }
    }
}

impl<T, C: Cleanup<T>> Drop for CoTaskMemArrayOut<T, C> {
    fn drop(&mut self) {
        if self.ptr.is_null() {
            return;
        }

        // SAFETY: `new`/`set_len` establish the initialized range and cleanup
        // contract before the output pointer is exposed to safe Rust again.
        unsafe {
            self.cleanup
                .as_mut()
                .expect("array output cleanup policy is always present")
                .cleanup(self.ptr, self.len);
            CoTaskMemFree(Some(self.ptr.cast()));
        }
        self.ptr = ptr::null_mut();
        self.len = 0;
    }
}

/// A null-initialized COM output pointer for one task allocation.
///
/// It releases an unclaimed allocation on drop, including when the COM call
/// returns an error. Use [`CoTaskMemArrayOut`] for arrays so their element
/// cleanup contract is attached before the call.
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
            // SAFETY: `CoTaskMemOut` is only for a single output allocation
            // governed by the COM task allocator. Arrays use
            // `CoTaskMemArrayOut` so nested resources are cleaned first.
            unsafe { CoTaskMemFree(Some(self.ptr.cast())) };
            self.ptr = ptr::null_mut();
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};

    use super::*;
    use crate::{CoTaskMemArrayBuilder, DropElements};

    struct CountDrop(Arc<AtomicUsize>);

    impl Drop for CountDrop {
        fn drop(&mut self) {
            self.0.fetch_add(1, Ordering::Relaxed);
        }
    }

    #[test]
    fn array_output_deep_cleans_when_not_adopted() {
        let drops = Arc::new(AtomicUsize::new(0));
        let mut builder = CoTaskMemArrayBuilder::new(2, DropElements).unwrap();
        builder.push(CountDrop(drops.clone())).ok().unwrap();
        builder.push(CountDrop(drops.clone())).ok().unwrap();
        let (ptr, len) = builder.finish().unwrap().into_raw_parts();

        let mut output = CoTaskMemArrayOut::<CountDrop, _>::new(len, DropElements);
        unsafe { output.as_mut_ptr().write(ptr) };
        drop(output);

        assert_eq!(drops.load(Ordering::Relaxed), 2);
    }
}
