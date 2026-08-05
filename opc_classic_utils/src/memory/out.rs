use std::marker::PhantomData;
use std::ptr;

use opc_classic_types::Result;
use windows::Win32::System::Com::CoTaskMemFree;

use super::{Cleanup, CoTaskMemArray, CoTaskMemObject, OwnedPwstr};

/// A null-initialized COM array output with its cleanup contract attached from
/// construction time.
///
/// Unlike [`CoTaskMemOut`], this guard knows how many initialized elements it
/// owns and how to release their nested resources. Dropping it before
/// [`Self::into_array`] therefore performs the same deep cleanup as the final
/// array owner.
pub struct CoTaskMemArrayOut<T, C: Cleanup<T>> {
    ptr: *mut T,
    /// `Some` means the number of initialized elements is trusted. `None`
    /// denotes a count reported by a foreign call that has not yet been
    /// validated against its HRESULT.
    len: Option<usize>,
    cleanup: Option<C>,
    _type: PhantomData<T>,
}

impl<T, C: Cleanup<T>> CoTaskMemArrayOut<T, C> {
    /// Creates an output guard for an array whose initialized length is known.
    pub const fn new(len: usize, cleanup: C) -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: Some(len),
            cleanup: Some(cleanup),
            _type: PhantomData,
        }
    }

    /// Creates an output guard whose element count is returned by the foreign
    /// call. The count remains untrusted until [`Self::commit_len`] is called
    /// after the call has returned a successful status.
    ///
    /// If the call fails, dropping this guard releases the outer task-memory
    /// allocation but does not inspect any nested elements using the untrusted
    /// count. This is intentional: a failing server is not allowed to make us
    /// walk an arbitrary number of pointers. For arrays with a known fixed
    /// length, use [`Self::new`] instead so nested resources are cleaned even on
    /// failure.
    pub const fn new_reported(cleanup: C) -> Self {
        Self {
            ptr: ptr::null_mut(),
            len: None,
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

    /// Commits the initialized length reported by the foreign call.
    ///
    /// # Safety
    ///
    /// If the output pointer is non-null, it must address at least `len`
    /// initialized values governed by this guard's cleanup policy. Call this
    /// only after checking that the foreign call succeeded (or after otherwise
    /// validating the count and allocation contract).
    pub unsafe fn commit_len(&mut self, len: usize) {
        self.len = Some(len);
    }

    /// Transfers the returned allocation into the final array owner.
    ///
    /// # Safety
    ///
    /// A non-null pointer must refer to `self.len` initialized values allocated
    /// with `CoTaskMemAlloc`, and the cleanup policy supplied to [`Self::new`]
    /// must match those values.
    pub unsafe fn into_array(mut self) -> Result<CoTaskMemArray<T, C>> {
        let Some(len) = self.len else {
            return Err(opc_classic_types::Error::invalid_argument(
                "reported task-memory array length was not committed",
            ));
        };
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        let cleanup = self
            .cleanup
            .take()
            .expect("array output cleanup policy is always present");
        unsafe { CoTaskMemArray::from_raw_parts(ptr, len, cleanup) }
    }

    /// Commits `len` and transfers the allocation in one operation.
    ///
    /// This is useful immediately after a successful foreign call and avoids
    /// accidentally adopting a reported count before the HRESULT has been
    /// checked.
    ///
    /// # Safety
    ///
    /// The same requirements as [`Self::commit_len`] and [`Self::into_array`]
    /// apply.
    pub unsafe fn into_array_with_len(mut self, len: usize) -> Result<CoTaskMemArray<T, C>> {
        unsafe { self.commit_len(len) };
        unsafe { self.into_array() }
    }
}

impl<T, C: Cleanup<T>> Drop for CoTaskMemArrayOut<T, C> {
    fn drop(&mut self) {
        if self.ptr.is_null() {
            return;
        }

        // SAFETY: `new` establishes a fixed initialized range, or
        // `commit_len` establishes a validated reported range. An uncommitted
        // reported count intentionally performs outer-only cleanup.
        unsafe {
            if let Some(len) = self.len {
                self.cleanup
                    .as_mut()
                    .expect("array output cleanup policy is always present")
                    .cleanup(self.ptr, len);
            }
            CoTaskMemFree(Some(self.ptr.cast()));
        }
        self.ptr = ptr::null_mut();
        self.len = Some(0);
    }
}

/// A single task-memory object with an explicit nested-element cleanup policy.
///
/// This is the object-shaped counterpart to [`CoTaskMemArray`]. It is useful
/// for COM methods returning `T**` where exactly one `T` is produced (for
/// example a status structure). The object is still allocated with
/// `CoTaskMemAlloc`, and its cleanup policy receives an initialized count of
/// one before the outer allocation is released.
pub struct CoTaskMemObjectOut<T, C: Cleanup<T>> {
    inner: CoTaskMemArrayOut<T, C>,
}

impl<T, C: Cleanup<T>> CoTaskMemObjectOut<T, C> {
    /// Creates a guard for one known returned object.
    pub const fn new(cleanup: C) -> Self {
        Self {
            inner: CoTaskMemArrayOut::new(1, cleanup),
        }
    }

    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        self.inner.as_mut_ptr()
    }

    pub fn is_null(&self) -> bool {
        self.inner.is_null()
    }

    /// Transfers the object allocation into its owning object wrapper.
    ///
    /// # Safety
    ///
    /// The foreign pointer must be null or point to one initialized `T` in a
    /// `CoTaskMemAlloc` allocation, and `cleanup` must match its nested values.
    pub unsafe fn into_object(self) -> Result<CoTaskMemObject<T, C>> {
        let array = unsafe { self.inner.into_array() }?;
        Ok(CoTaskMemObject::from_array(array))
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
    use crate::{CoTaskMemArrayBuilder, DropElements, NoCleanup};

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

    #[test]
    fn reported_array_output_cleans_after_committed_transfer() {
        let drops = Arc::new(AtomicUsize::new(0));
        let mut builder = CoTaskMemArrayBuilder::new(2, DropElements).unwrap();
        builder.push(CountDrop(drops.clone())).ok().unwrap();
        builder.push(CountDrop(drops.clone())).ok().unwrap();
        let (ptr, len) = builder.finish().unwrap().into_raw_parts();

        let mut output = CoTaskMemArrayOut::<CountDrop, _>::new_reported(DropElements);
        unsafe { output.as_mut_ptr().write(ptr) };
        unsafe { output.commit_len(len) };
        let array = unsafe { output.into_array() }.unwrap();
        assert_eq!(drops.load(Ordering::Relaxed), 0);

        drop(array);
        assert_eq!(drops.load(Ordering::Relaxed), 2);
    }

    struct RecordCleanup(Arc<AtomicUsize>);

    // SAFETY: This test policy only records the validated initialized length
    // and never reads or releases the plain integer elements.
    unsafe impl Cleanup<u32> for RecordCleanup {
        unsafe fn cleanup(&mut self, _ptr: *mut u32, initialized: usize) {
            self.0.store(initialized, Ordering::Relaxed);
        }
    }

    #[test]
    fn uncommitted_reported_length_is_never_used_for_cleanup() {
        let mut builder = CoTaskMemArrayBuilder::new(2, NoCleanup).unwrap();
        builder.push(10u32).ok().unwrap();
        builder.push(20u32).ok().unwrap();
        let (ptr, _) = builder.finish().unwrap().into_raw_parts();
        let cleaned = Arc::new(AtomicUsize::new(usize::MAX));

        let mut output = CoTaskMemArrayOut::new_reported(RecordCleanup(cleaned.clone()));
        unsafe { output.as_mut_ptr().write(ptr) };
        let error = unsafe { output.into_array() }.err().unwrap();

        assert_eq!(
            error.message(),
            "reported task-memory array length was not committed"
        );
        assert_eq!(cleaned.load(Ordering::Relaxed), usize::MAX);
    }

    #[test]
    fn object_output_uses_single_element_cleanup() {
        let drops = Arc::new(AtomicUsize::new(0));
        let mut builder = CoTaskMemArrayBuilder::new(1, DropElements).unwrap();
        builder.push(CountDrop(drops.clone())).ok().unwrap();
        let (ptr, len) = builder.finish().unwrap().into_raw_parts();
        assert_eq!(len, 1);

        let mut output = CoTaskMemObjectOut::<CountDrop, _>::new(DropElements);
        unsafe { output.as_mut_ptr().write(ptr) };
        let object = unsafe { output.into_object() }.unwrap();
        assert_eq!(drops.load(Ordering::Relaxed), 0);

        drop(object);
        assert_eq!(drops.load(Ordering::Relaxed), 1);
    }
}
