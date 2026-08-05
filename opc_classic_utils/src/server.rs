use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;
use std::sync::atomic::{AtomicU32, Ordering};

use opc_classic_types::{ClassContext, ComObject, Error, ErrorCode, Guid, Result};
use windows::Win32::Foundation::{CLASS_E_NOAGGREGATION, E_POINTER, E_UNEXPECTED};
use windows::Win32::System::Com::{
    CLSCTX, CoRegisterClassObject, CoRevokeClassObject, IClassFactory, IClassFactory_Impl,
    REGCLS_MULTIPLEUSE,
};
use windows_core::{Error as AbiError, GUID, IUnknown, Ref, Result as AbiResult};

use crate::ComApartment;

/// Prevents Rust unwinding from crossing a foreign-function boundary.
pub fn catch_ffi<T>(f: impl FnOnce() -> Result<T>) -> Result<T> {
    catch_unwind(AssertUnwindSafe(f)).unwrap_or_else(|_| {
        Err(Error::new(
            opc_classic_types::ErrorKind::Panic,
            ErrorCode::UNEXPECTED,
            "panic in OPC callback",
        ))
    })
}

/// Runs an ABI closure without allowing a panic to cross COM.
pub fn catch_abi<T, E>(
    f: impl FnOnce() -> core::result::Result<T, E>,
    panic_error: E,
) -> core::result::Result<T, E> {
    catch_unwind(AssertUnwindSafe(f)).unwrap_or(Err(panic_error))
}

/// Validates and borrows a foreign input array for the current call.
///
/// # Safety
///
/// A non-zero length requires `ptr` to point to that many initialized values,
/// and the returned slice must not outlive the invocation.
pub unsafe fn borrow_input<'call, T>(ptr: *const T, len: u32) -> Result<&'call [T]> {
    if len == 0 {
        return Ok(&[]);
    }
    if ptr.is_null() {
        return Err(Error::null_pointer("null OPC input array"));
    }
    Ok(unsafe { std::slice::from_raw_parts(ptr, len as usize) })
}

/// Validates and zero-initializes a required foreign output pointer.
///
/// # Safety
///
/// `out` must be writable when non-null and valid for one `T`.
pub unsafe fn initialize_output<T: Default>(out: *mut T) -> Result<()> {
    if out.is_null() {
        return Err(Error::null_pointer("null OPC output pointer"));
    }
    unsafe { out.write(T::default()) };
    Ok(())
}

type Activator = dyn Fn() -> Result<ComObject> + Send + Sync + 'static;

#[windows_core::implement(IClassFactory)]
struct ClassFactoryAdapter {
    activate: Arc<Activator>,
    server_locks: Arc<AtomicU32>,
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IClassFactory_Impl for ClassFactoryAdapter_Impl {
    fn CreateInstance(
        &self,
        outer: Ref<IUnknown>,
        iid: *const GUID,
        output: *mut *mut core::ffi::c_void,
    ) -> AbiResult<()> {
        catch_abi(
            || {
                if output.is_null() || iid.is_null() {
                    return Err(AbiError::from_hresult(E_POINTER));
                }
                unsafe { output.write(core::ptr::null_mut()) };
                if !outer.is_null() {
                    return Err(AbiError::from_hresult(CLASS_E_NOAGGREGATION));
                }
                let object = (self.activate)().map_err(to_abi_error)?;
                let iid = unsafe { &*iid };
                let iid = Guid::new(iid.data1, iid.data2, iid.data3, iid.data4);
                let requested = object.query_interface(&iid).map_err(to_abi_error)?;
                unsafe { output.write(requested.into_raw()) };
                Ok(())
            },
            AbiError::from_hresult(E_UNEXPECTED),
        )
    }

    fn LockServer(&self, lock: windows_core::BOOL) -> AbiResult<()> {
        catch_abi(
            || {
                let update = if lock.as_bool() {
                    self.server_locks
                        .fetch_update(Ordering::AcqRel, Ordering::Acquire, |count| {
                            count.checked_add(1)
                        })
                } else {
                    self.server_locks
                        .fetch_update(Ordering::AcqRel, Ordering::Acquire, |count| {
                            count.checked_sub(1)
                        })
                };
                update
                    .map(|_| ())
                    .map_err(|_| AbiError::from_hresult(E_UNEXPECTED))
            },
            AbiError::from_hresult(E_UNEXPECTED),
        )
    }
}

/// A Rust-backed COM class factory whose public API contains only project-owned
/// and standard-library types.
pub struct ClassFactory {
    inner: IClassFactory,
    server_locks: Arc<AtomicU32>,
}

impl ClassFactory {
    pub fn new(activate: impl Fn() -> Result<ComObject> + Send + Sync + 'static) -> Self {
        let server_locks = Arc::new(AtomicU32::new(0));
        let adapter = ClassFactoryAdapter {
            activate: Arc::new(activate),
            server_locks: server_locks.clone(),
        };
        Self {
            inner: adapter.into(),
            server_locks,
        }
    }

    pub fn server_lock_count(&self) -> u32 {
        self.server_locks.load(Ordering::Acquire)
    }
}

/// A class-object registration revoked on drop and bound to its COM apartment.
pub struct LocalClassRegistration<'apartment> {
    cookie: Option<u32>,
    _factory: IClassFactory,
    _apartment: &'apartment ComApartment,
}

impl<'apartment> LocalClassRegistration<'apartment> {
    pub fn register(
        apartment: &'apartment ComApartment,
        class_id: &Guid,
        factory: ClassFactory,
    ) -> Result<Self> {
        Self::register_in_context(apartment, class_id, factory, ClassContext::LOCAL_SERVER)
    }

    /// Registers the class factory with `REGCLS_MULTIPLEUSE` in the requested
    /// activation context.
    ///
    /// Other registration policies are intentionally not exposed until they
    /// can be represented by a project-owned type rather than a Windows enum.
    pub fn register_in_context(
        apartment: &'apartment ComApartment,
        class_id: &Guid,
        factory: ClassFactory,
        context: ClassContext,
    ) -> Result<Self> {
        let class_id = GUID::from_values(
            class_id.data1,
            class_id.data2,
            class_id.data3,
            class_id.data4,
        );
        let cookie = unsafe {
            CoRegisterClassObject(
                &class_id,
                &factory.inner,
                CLSCTX(context.bits()),
                REGCLS_MULTIPLEUSE,
            )
        }
        .map_err(from_abi_error)?;
        Ok(Self {
            cookie: Some(cookie),
            _factory: factory.inner,
            _apartment: apartment,
        })
    }

    pub fn revoke(mut self) -> Result<()> {
        if let Some(cookie) = self.cookie.take() {
            unsafe { CoRevokeClassObject(cookie) }.map_err(from_abi_error)
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

fn to_abi_error(error: Error) -> AbiError {
    let status = error.boundary_status();
    let status = windows_core::HRESULT(status.raw());
    if status.0 == ErrorCode::PARTIAL_SUCCESS.raw() {
        AbiError::from_hresult(status)
    } else {
        AbiError::new(status, error.message())
    }
}

fn from_abi_error(error: AbiError) -> Error {
    Error::from_code(ErrorCode::from_raw(error.code().0)).with_message(error.message())
}
