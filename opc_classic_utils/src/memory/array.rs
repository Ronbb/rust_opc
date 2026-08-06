use std::marker::PhantomData;
use std::mem::{ManuallyDrop, MaybeUninit};
use std::ptr::{self, NonNull};

use opc_classic_types::{Error, Result};
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};

/// Describes how initialized elements inside a `CoTaskMem` allocation are released.
///
/// # Safety
///
/// Implementations must accept exactly `initialized` valid consecutive values at
/// `ptr`. They must release resources owned by those values without freeing the
/// outer allocation itself.
pub unsafe trait Cleanup<T> {
    /// Releases resources owned by an initialized prefix without freeing the
    /// outer allocation.
    ///
    /// # Safety
    ///
    /// `ptr` must address at least `initialized` consecutive, initialized `T`
    /// values, each exclusively owned by the caller.
    unsafe fn cleanup(&mut self, ptr: *mut T, initialized: usize);
}

/// Leaves elements untouched and only releases the outer `CoTaskMem` allocation.
///
/// This is appropriate for plain ABI values and raw structures whose nested
/// resources are handled separately by a method-specific decoder.
#[derive(Clone, Copy, Debug, Default)]
pub struct NoCleanup;

// SAFETY: This policy intentionally performs no element operation.
unsafe impl<T> Cleanup<T> for NoCleanup {
    unsafe fn cleanup(&mut self, _ptr: *mut T, _initialized: usize) {}
}

/// Runs Rust drop glue for every initialized element before releasing the outer
/// allocation. This is useful for `VARIANT` and structures whose Rust drop glue
/// correctly matches their COM cleanup contract.
#[derive(Clone, Copy, Debug, Default)]
pub struct DropElements;

// SAFETY: Every element is dropped exactly once and the outer allocation is left
// for `CoTaskMemArray` to release.
unsafe impl<T> Cleanup<T> for DropElements {
    unsafe fn cleanup(&mut self, ptr: *mut T, initialized: usize) {
        for index in 0..initialized {
            unsafe { ptr::drop_in_place(ptr.add(index)) };
        }
    }
}

/// Releases every `PWSTR` element with `CoTaskMemFree` before releasing the
/// pointer array itself.
#[derive(Clone, Copy, Debug, Default)]
pub struct FreePwstrElements;

// SAFETY: OPC string arrays use one task allocation per non-null string and one
// task allocation for the pointer array.
unsafe impl Cleanup<*mut u16> for FreePwstrElements {
    unsafe fn cleanup(&mut self, ptr: *mut *mut u16, initialized: usize) {
        for index in 0..initialized {
            let value = unsafe { *ptr.add(index) };
            if !value.is_null() {
                unsafe { CoTaskMemFree(Some(value.cast())) };
            }
        }
    }
}

/// Owns a contiguous, initialized array allocated with `CoTaskMemAlloc`.
///
/// `C` makes element cleanup explicit. The outer allocation is always released
/// with `CoTaskMemFree`.
pub struct CoTaskMemArray<T, C: Cleanup<T>> {
    ptr: Option<NonNull<T>>,
    len: usize,
    cleanup: C,
    _owns: PhantomData<T>,
}

impl<T, C: Cleanup<T>> CoTaskMemArray<T, C> {
    /// Adopts an array returned by COM.
    ///
    /// # Safety
    ///
    /// For non-zero `len`, `ptr` must point to `len` initialized `T` values in a
    /// single allocation obtained from `CoTaskMemAlloc`. `cleanup` must match the
    /// ownership contract of every element.
    pub unsafe fn from_raw_parts(ptr: *mut T, len: usize, cleanup: C) -> Result<Self> {
        if len != 0 && ptr.is_null() {
            return Err(Error::null_pointer("null task-memory array"));
        }

        Ok(Self {
            ptr: NonNull::new(ptr),
            len,
            cleanup,
            _owns: PhantomData,
        })
    }

    pub fn len(&self) -> usize {
        self.len
    }

    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    pub fn as_ptr(&self) -> *const T {
        self.ptr.map_or(ptr::null(), |ptr| ptr.as_ptr())
    }

    pub fn as_mut_ptr(&mut self) -> *mut T {
        self.ptr.map_or(ptr::null_mut(), |ptr| ptr.as_ptr())
    }

    pub fn as_slice(&self) -> &[T] {
        if let Some(ptr) = self.ptr {
            // SAFETY: Construction guarantees `len` initialized elements.
            unsafe { std::slice::from_raw_parts(ptr.as_ptr(), self.len) }
        } else {
            &[]
        }
    }

    pub fn as_mut_slice(&mut self) -> &mut [T] {
        if let Some(ptr) = self.ptr {
            // SAFETY: The allocation is uniquely owned by `self`.
            unsafe { std::slice::from_raw_parts_mut(ptr.as_ptr(), self.len) }
        } else {
            &mut []
        }
    }

    /// Transfers the allocation and its element ownership to the caller.
    pub fn into_raw_parts(self) -> (*mut T, usize) {
        let mut this = ManuallyDrop::new(self);
        let raw = this.ptr.map_or(ptr::null_mut(), |ptr| ptr.as_ptr());
        let len = this.len;
        // The cleanup policy itself is not transferred with the allocation.
        // Drop its Rust-owned state without running element cleanup.
        unsafe { ptr::drop_in_place(&mut this.cleanup) };
        (raw, len)
    }
}

impl<T, C: Cleanup<T>> Drop for CoTaskMemArray<T, C> {
    fn drop(&mut self) {
        let Some(ptr) = self.ptr.take() else {
            return;
        };

        // SAFETY: The constructor establishes the initialized range and cleanup
        // contract. The allocation is uniquely owned here.
        unsafe {
            self.cleanup.cleanup(ptr.as_ptr(), self.len);
            CoTaskMemFree(Some(ptr.as_ptr().cast()));
        }
        self.len = 0;
    }
}

/// Transactionally builds an initialized `CoTaskMem` array.
///
/// If construction fails or unwinds, only the initialized prefix is cleaned.
pub struct CoTaskMemArrayBuilder<T, C: Cleanup<T>> {
    ptr: Option<NonNull<MaybeUninit<T>>>,
    capacity: usize,
    initialized: usize,
    cleanup: C,
}

impl<T, C: Cleanup<T>> CoTaskMemArrayBuilder<T, C> {
    pub fn new(capacity: usize, cleanup: C) -> Result<Self> {
        if capacity == 0 {
            return Ok(Self {
                ptr: None,
                capacity: 0,
                initialized: 0,
                cleanup,
            });
        }

        let bytes = std::mem::size_of::<T>()
            .checked_mul(capacity)
            .filter(|bytes| *bytes != 0)
            .ok_or_else(|| Error::invalid_argument("invalid task-memory array size"))?;
        let ptr = unsafe { CoTaskMemAlloc(bytes) }.cast::<MaybeUninit<T>>();
        let ptr = NonNull::new(ptr)
            .ok_or_else(|| Error::out_of_memory("task-memory allocation failed"))?;
        if ptr.as_ptr().addr() % std::mem::align_of::<T>() != 0 {
            unsafe { CoTaskMemFree(Some(ptr.as_ptr().cast())) };
            return Err(Error::invalid_argument("task-memory alignment mismatch"));
        }

        Ok(Self {
            ptr: Some(ptr),
            capacity,
            initialized: 0,
            cleanup,
        })
    }

    pub fn push(&mut self, value: T) -> core::result::Result<(), T> {
        if self.initialized == self.capacity {
            return Err(value);
        }

        let ptr = self.ptr.expect("non-empty builder must have an allocation");
        // SAFETY: `initialized < capacity` and this slot has not been written.
        unsafe {
            ptr.as_ptr()
                .add(self.initialized)
                .write(MaybeUninit::new(value))
        };
        self.initialized += 1;
        Ok(())
    }

    pub fn finish(self) -> Result<CoTaskMemArray<T, C>> {
        if self.initialized != self.capacity {
            return Err(Error::invalid_argument("task-memory array is incomplete"));
        }

        let this = ManuallyDrop::new(self);
        // SAFETY: `cleanup` is moved out and `this` will not be dropped.
        let cleanup = unsafe { ptr::read(&this.cleanup) };
        Ok(CoTaskMemArray {
            ptr: this.ptr.map(|ptr| ptr.cast()),
            len: this.capacity,
            cleanup,
            _owns: PhantomData,
        })
    }
}

impl<T, C: Cleanup<T>> Drop for CoTaskMemArrayBuilder<T, C> {
    fn drop(&mut self) {
        let Some(ptr) = self.ptr.take() else {
            return;
        };

        // SAFETY: Only the initialized prefix contains valid `T` values.
        unsafe {
            self.cleanup.cleanup(ptr.as_ptr().cast(), self.initialized);
            CoTaskMemFree(Some(ptr.as_ptr().cast()));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};

    struct CountDrop {
        drops: Arc<AtomicUsize>,
    }

    impl Drop for CountDrop {
        fn drop(&mut self) {
            self.drops.fetch_add(1, Ordering::Relaxed);
        }
    }

    #[test]
    fn builder_rolls_back_initialized_prefix() {
        let drops = Arc::new(AtomicUsize::new(0));
        {
            let mut builder = CoTaskMemArrayBuilder::new(3, DropElements).unwrap();
            builder
                .push(CountDrop {
                    drops: drops.clone(),
                })
                .ok()
                .unwrap();
            builder
                .push(CountDrop {
                    drops: drops.clone(),
                })
                .ok()
                .unwrap();
        }
        assert_eq!(drops.load(Ordering::Relaxed), 2);
    }

    #[test]
    fn finished_array_drops_all_elements() {
        let drops = Arc::new(AtomicUsize::new(0));
        {
            let mut builder = CoTaskMemArrayBuilder::new(2, DropElements).unwrap();
            builder
                .push(CountDrop {
                    drops: drops.clone(),
                })
                .ok()
                .unwrap();
            builder
                .push(CountDrop {
                    drops: drops.clone(),
                })
                .ok()
                .unwrap();
            let array = builder.finish().unwrap();
            assert_eq!(array.len(), 2);
        }
        assert_eq!(drops.load(Ordering::Relaxed), 2);
    }

    #[test]
    fn zero_length_array_is_supported() {
        let builder = CoTaskMemArrayBuilder::<u32, _>::new(0, NoCleanup).unwrap();
        let array = builder.finish().unwrap();
        assert!(array.is_empty());
        assert!(array.as_ptr().is_null());
    }

    #[test]
    fn raw_transfer_does_not_clean_initialized_elements() {
        let drops = Arc::new(AtomicUsize::new(0));
        let mut builder = CoTaskMemArrayBuilder::new(2, DropElements).unwrap();
        builder
            .push(CountDrop {
                drops: drops.clone(),
            })
            .ok()
            .unwrap();
        builder
            .push(CountDrop {
                drops: drops.clone(),
            })
            .ok()
            .unwrap();

        let (ptr, len) = builder.finish().unwrap().into_raw_parts();
        assert_eq!(drops.load(Ordering::Relaxed), 0);

        // SAFETY: ownership was transferred by `into_raw_parts` immediately
        // above and is re-adopted exactly once with the same cleanup policy.
        let array = unsafe { CoTaskMemArray::from_raw_parts(ptr, len, DropElements) }.unwrap();
        drop(array);
        assert_eq!(drops.load(Ordering::Relaxed), 2);
    }

    #[test]
    fn push_past_capacity_returns_the_original_value() {
        let stored_drops = Arc::new(AtomicUsize::new(0));
        let returned_drops = Arc::new(AtomicUsize::new(0));
        let mut builder = CoTaskMemArrayBuilder::new(1, DropElements).unwrap();
        builder
            .push(CountDrop {
                drops: stored_drops.clone(),
            })
            .ok()
            .unwrap();

        let returned = builder
            .push(CountDrop {
                drops: returned_drops.clone(),
            })
            .expect_err("a full builder returns the unconsumed value");
        assert_eq!(stored_drops.load(Ordering::Relaxed), 0);
        assert_eq!(returned_drops.load(Ordering::Relaxed), 0);

        drop(returned);
        assert_eq!(returned_drops.load(Ordering::Relaxed), 1);
        drop(builder);
        assert_eq!(stored_drops.load(Ordering::Relaxed), 1);
    }
}
