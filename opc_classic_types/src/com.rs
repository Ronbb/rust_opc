use std::ffi::c_void;
use std::fmt;
use std::marker::PhantomData;
use std::ptr::{self, NonNull};
use std::rc::Rc;

use crate::{Error, ErrorCode, Result};

/// A GUID with the same field layout as the COM ABI, defined without a Windows
/// crate dependency.
#[repr(C)]
#[derive(Clone, Copy, Default, Eq, Hash, PartialEq)]
pub struct Guid {
    pub data1: u32,
    pub data2: u16,
    pub data3: u16,
    pub data4: [u8; 8],
}

impl Guid {
    pub const ZERO: Self = Self {
        data1: 0,
        data2: 0,
        data3: 0,
        data4: [0; 8],
    };

    pub const fn new(data1: u32, data2: u16, data3: u16, data4: [u8; 8]) -> Self {
        Self {
            data1,
            data2,
            data3,
            data4,
        }
    }
}

impl fmt::Debug for Guid {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{:08x}-{:04x}-{:04x}-{:02x}{:02x}-",
            self.data1, self.data2, self.data3, self.data4[0], self.data4[1]
        )?;
        for value in &self.data4[2..] {
            write!(f, "{value:02x}")?;
        }
        Ok(())
    }
}

/// COM activation contexts represented as ordinary bit flags.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct ClassContext(u32);

impl ClassContext {
    pub const IN_PROCESS_SERVER: Self = Self(0x1);
    pub const IN_PROCESS_HANDLER: Self = Self(0x2);
    pub const LOCAL_SERVER: Self = Self(0x4);
    pub const REMOTE_SERVER: Self = Self(0x10);
    pub const ALL: Self = Self(0x17);

    pub const fn from_bits(bits: u32) -> Self {
        Self(bits)
    }

    pub const fn bits(self) -> u32 {
        self.0
    }

    pub const fn contains(self, other: Self) -> bool {
        self.0 & other.0 == other.0
    }
}

impl core::ops::BitOr for ClassContext {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

/// A 100-nanosecond timestamp since 1601-01-01 UTC, matching OPC Classic's
/// timestamp precision without exposing `FILETIME`.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Default, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Timestamp(u64);

impl Timestamp {
    pub const UNIX_EPOCH_TICKS: u64 = 116_444_736_000_000_000;

    pub const fn from_ticks(ticks: u64) -> Self {
        Self(ticks)
    }

    pub const fn ticks(self) -> u64 {
        self.0
    }

    pub fn from_system_time(value: std::time::SystemTime) -> Result<Self> {
        let duration = value
            .duration_since(std::time::UNIX_EPOCH)
            .map_err(|_| Error::invalid_argument("timestamp is before the Unix epoch"))?;
        let ticks = duration
            .as_secs()
            .checked_mul(10_000_000)
            .and_then(|ticks| ticks.checked_add(u64::from(duration.subsec_nanos() / 100)))
            .and_then(|ticks| ticks.checked_add(Self::UNIX_EPOCH_TICKS))
            .ok_or_else(|| Error::invalid_argument("timestamp is out of range"))?;
        Ok(Self(ticks))
    }

    pub fn to_system_time(self) -> Result<std::time::SystemTime> {
        let ticks = self
            .0
            .checked_sub(Self::UNIX_EPOCH_TICKS)
            .ok_or_else(|| Error::invalid_argument("timestamp predates the Unix epoch"))?;
        Ok(std::time::UNIX_EPOCH
            + std::time::Duration::new(ticks / 10_000_000, (ticks % 10_000_000) as u32 * 100))
    }
}

#[repr(C)]
struct UnknownVtable {
    query_interface: unsafe extern "system" fn(*mut c_void, *const Guid, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
}

/// An owning, apartment-bound COM identity pointer with no Windows crate in its
/// public API.
pub struct ComObject {
    ptr: NonNull<c_void>,
    _apartment_bound: PhantomData<Rc<()>>,
}

impl ComObject {
    /// Adopts one owned COM reference.
    ///
    /// # Safety
    ///
    /// `ptr` must be a non-null COM interface pointer whose first three vtable
    /// entries are `IUnknown`, and ownership of exactly one reference transfers
    /// to the returned value.
    pub unsafe fn from_raw_owned(ptr: *mut c_void) -> Result<Self> {
        let ptr = NonNull::new(ptr).ok_or_else(|| Error::null_pointer("null COM object"))?;
        Ok(Self {
            ptr,
            _apartment_bound: PhantomData,
        })
    }

    pub fn as_ptr(&self) -> *mut c_void {
        self.ptr.as_ptr()
    }

    pub fn query_interface(&self, interface_id: &Guid) -> Result<Self> {
        let mut output = ptr::null_mut();
        let code = unsafe {
            (self.vtable().query_interface)(self.ptr.as_ptr(), interface_id, &mut output)
        };
        let code = ErrorCode::from_raw(code);
        if code.is_failure() {
            return Err(Error::from_code(code).with_message("COM interface is unavailable"));
        }
        unsafe { Self::from_raw_owned(output) }
    }

    pub fn into_raw(self) -> *mut c_void {
        let ptr = self.ptr.as_ptr();
        core::mem::forget(self);
        ptr
    }

    fn vtable(&self) -> &UnknownVtable {
        unsafe { &**self.ptr.as_ptr().cast::<*const UnknownVtable>() }
    }
}

impl Clone for ComObject {
    fn clone(&self) -> Self {
        unsafe { (self.vtable().add_ref)(self.ptr.as_ptr()) };
        Self {
            ptr: self.ptr,
            _apartment_bound: PhantomData,
        }
    }
}

impl fmt::Debug for ComObject {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_tuple("ComObject").field(&self.ptr).finish()
    }
}

impl Drop for ComObject {
    fn drop(&mut self) {
        unsafe { (self.vtable().release)(self.ptr.as_ptr()) };
    }
}
