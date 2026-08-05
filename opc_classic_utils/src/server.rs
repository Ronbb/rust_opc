use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::{CLASS_E_NOAGGREGATION, E_POINTER, E_UNEXPECTED};
use windows::Win32::System::Com::{
    CLSCTX, CLSCTX_LOCAL_SERVER, CoRegisterClassObject, CoRevokeClassObject, IClassFactory,
    IClassFactory_Impl, REGCLS, REGCLS_MULTIPLEUSE,
};
use windows_core::{Error, GUID, IUnknown, Interface, Ref, Result};

use crate::ComApartment;

/// Prevents Rust unwinding from crossing a COM ABI boundary.
pub fn catch_ffi<T>(f: impl FnOnce() -> Result<T>) -> Result<T> {
    catch_unwind(AssertUnwindSafe(f)).unwrap_or_else(|_| Err(Error::from_hresult(E_UNEXPECTED)))
}

/// Validates and borrows a COM input array for the duration of the current call.
///
/// # Safety
///
/// A non-zero length requires `ptr` to point to that many initialized values,
/// and the returned slice must not outlive the COM method invocation.
pub unsafe fn borrow_input<'call, T>(ptr: *const T, len: u32) -> Result<&'call [T]> {
    if len == 0 {
        return Ok(&[]);
    }
    if ptr.is_null() {
        return Err(Error::from_hresult(windows::Win32::Foundation::E_POINTER));
    }
    Ok(unsafe { std::slice::from_raw_parts(ptr, len as usize) })
}

/// Validates a required COM output pointer and initializes it with `Default`.
///
/// # Safety
///
/// `out` must be writable when non-null and valid for a single `T`.
pub unsafe fn initialize_output<T: Default>(out: *mut T) -> Result<()> {
    if out.is_null() {
        return Err(Error::from_hresult(windows::Win32::Foundation::E_POINTER));
    }
    unsafe { out.write(T::default()) };
    Ok(())
}

type Activator = dyn Fn() -> Result<IUnknown> + Send + Sync + 'static;

/// A Rust-backed COM class factory.
///
/// The activator returns the identity `IUnknown` for a fresh object. The factory
/// performs aggregation checks and `QueryInterface` for the requested interface.
#[windows_core::implement(IClassFactory)]
pub struct ClassFactory {
    activate: Arc<Activator>,
    server_locks: Arc<AtomicU32>,
}

impl ClassFactory {
    pub fn new(activate: impl Fn() -> Result<IUnknown> + Send + Sync + 'static) -> Self {
        Self {
            activate: Arc::new(activate),
            server_locks: Arc::new(AtomicU32::new(0)),
        }
    }

    pub fn server_lock_count(&self) -> u32 {
        self.server_locks.load(Ordering::Acquire)
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IClassFactory_Impl for ClassFactory_Impl {
    fn CreateInstance(
        &self,
        outer: Ref<IUnknown>,
        iid: *const GUID,
        output: *mut *mut core::ffi::c_void,
    ) -> Result<()> {
        catch_ffi(|| {
            unsafe { initialize_output(output)? };
            if !outer.is_null() {
                return Err(Error::from_hresult(CLASS_E_NOAGGREGATION));
            }
            let iid = unsafe { iid.as_ref() }.ok_or_else(|| Error::from_hresult(E_POINTER))?;
            let object = (self.activate)()?;
            unsafe { object.query(iid, output) }.ok()
        })
    }

    fn LockServer(&self, lock: windows_core::BOOL) -> Result<()> {
        catch_ffi(|| {
            if lock.as_bool() {
                return self
                    .server_locks
                    .fetch_update(Ordering::AcqRel, Ordering::Acquire, |count| {
                        count.checked_add(1)
                    })
                    .map(|_| ())
                    .map_err(|_| Error::from_hresult(E_UNEXPECTED));
            }

            self.server_locks
                .fetch_update(Ordering::AcqRel, Ordering::Acquire, |count| {
                    count.checked_sub(1)
                })
                .map(|_| ())
                .map_err(|_| Error::from_hresult(E_UNEXPECTED))
        })
    }
}

/// A class-object registration that is revoked when dropped.
///
/// Borrowing the apartment makes the registration thread-bound and prevents COM
/// from being uninitialized before `CoRevokeClassObject` runs.
pub struct LocalClassRegistration<'apartment> {
    cookie: Option<u32>,
    _factory: IClassFactory,
    _apartment: &'apartment ComApartment,
}

impl<'apartment> LocalClassRegistration<'apartment> {
    pub fn register(
        apartment: &'apartment ComApartment,
        class_id: &GUID,
        factory: ClassFactory,
    ) -> Result<Self> {
        Self::register_with(
            apartment,
            class_id,
            factory.into(),
            CLSCTX_LOCAL_SERVER,
            REGCLS_MULTIPLEUSE,
        )
    }

    pub fn register_with(
        apartment: &'apartment ComApartment,
        class_id: &GUID,
        factory: IClassFactory,
        context: CLSCTX,
        flags: REGCLS,
    ) -> Result<Self> {
        let cookie = unsafe { CoRegisterClassObject(class_id, &factory, context, flags) }?;
        Ok(Self {
            cookie: Some(cookie),
            _factory: factory,
            _apartment: apartment,
        })
    }

    pub fn revoke(mut self) -> Result<()> {
        if let Some(cookie) = self.cookie.take() {
            unsafe { CoRevokeClassObject(cookie) }
        } else {
            Ok(())
        }
    }
}

impl Drop for LocalClassRegistration<'_> {
    fn drop(&mut self) {
        if let Some(cookie) = self.cookie.take() {
            let _ = unsafe { CoRevokeClassObject(cookie) };
        }
    }
}
