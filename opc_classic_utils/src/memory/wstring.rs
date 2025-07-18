use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::ptr;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows::core::PCWSTR;

/// A smart pointer for wide string pointers that the **caller allocates and callee frees**
///
/// This is used for input string parameters where the caller allocates memory
/// and the callee (COM function) is responsible for freeing it.
/// This wrapper does NOT free the memory when dropped.
#[repr(transparent)]
#[derive(Debug)]
pub struct CallerAllocatedWString {
    ptr: *mut u16,
}

impl CallerAllocatedWString {
    /// Creates a new `CallerAllocatedWString` from a raw wide string pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid wide string pointer
    /// allocated by the caller and that the callee will be responsible for freeing it.
    pub unsafe fn new(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CallerAllocatedWString` from a raw pointer, taking ownership
    pub fn from_raw(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CallerAllocatedWString` from a `PCWSTR`
    pub fn from_pcwstr(pcwstr: PCWSTR) -> Self {
        Self {
            ptr: pcwstr.as_ptr() as *mut u16,
        }
    }

    /// Allocates memory using `CoTaskMemAlloc` and creates a `CallerAllocatedWString`
    ///
    /// This allocates memory for a wide string that will be freed by the callee.
    pub fn allocate(len: usize) -> Result<Self, windows::core::Error> {
        let size = (len + 1) * std::mem::size_of::<u16>(); // +1 for null terminator
        let ptr = unsafe { CoTaskMemAlloc(size) };
        if ptr.is_null() {
            return Err(windows::core::Error::from_win32());
        }
        Ok(unsafe { Self::new(ptr.cast()) })
    }

    /// Creates a `CallerAllocatedWString` from a Rust string
    pub fn from_string(s: String) -> Result<Self, windows::core::Error> {
        use std::str::FromStr;
        Self::from_str(&s)
    }

    /// Creates a `CallerAllocatedWString` from an `OsStr`
    pub fn from_os_str(os_str: &OsStr) -> Result<Self, windows::core::Error> {
        let wide_string: Vec<u16> = os_str.encode_wide().chain(std::iter::once(0)).collect();
        let len = wide_string.len() - 1; // Exclude null terminator for allocation

        let ptr = Self::allocate(len)?;
        unsafe {
            std::ptr::copy_nonoverlapping(wide_string.as_ptr(), ptr.as_ptr(), wide_string.len());
        }
        Ok(ptr)
    }

    /// Converts the wide string to a Rust string slice
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_string(&self) -> Option<String> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        let os_string = OsString::from_wide(slice);
        Some(os_string.to_string_lossy().into_owned())
    }

    /// Converts the wide string to an `OsString`
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_os_string(&self) -> Option<OsString> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        Some(OsString::from_wide(slice))
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut u16 {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    pub fn into_raw(mut self) -> *mut u16 {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Converts to a `PCWSTR` for use with Windows APIs
    pub fn as_pcwstr(&self) -> PCWSTR {
        PCWSTR(self.ptr)
    }
}

impl Drop for CallerAllocatedWString {
    fn drop(&mut self) {
        // Do NOT free the memory - the callee is responsible for this
        self.ptr = ptr::null_mut();
    }
}

impl Default for CallerAllocatedWString {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl Clone for CallerAllocatedWString {
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

impl std::str::FromStr for CallerAllocatedWString {
    type Err = windows::core::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let wide_string: Vec<u16> = OsStr::new(s)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        let len = wide_string.len() - 1; // Exclude null terminator for allocation

        let ptr = Self::allocate(len)?;
        unsafe {
            std::ptr::copy_nonoverlapping(wide_string.as_ptr(), ptr.as_ptr(), wide_string.len());
        }
        Ok(ptr)
    }
}

/// A smart pointer for wide string pointers that the **callee allocates and caller frees**
///
/// This is used for output string parameters where the callee allocates memory
/// and the caller is responsible for freeing it using `CoTaskMemFree`.
#[repr(transparent)]
#[derive(Debug)]
pub struct CalleeAllocatedWString {
    ptr: *mut u16,
}

impl CalleeAllocatedWString {
    /// Creates a new `CalleeAllocatedWString` from a raw wide string pointer
    ///
    /// # Safety
    ///
    /// The caller must ensure that `ptr` is a valid wide string pointer
    /// allocated by the callee and that it will be freed using `CoTaskMemFree`.
    pub unsafe fn new(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CalleeAllocatedWString` from a raw pointer, taking ownership
    pub fn from_raw(ptr: *mut u16) -> Self {
        Self { ptr }
    }

    /// Creates a new `CalleeAllocatedWString` from a `PCWSTR`
    pub fn from_pcwstr(pcwstr: PCWSTR) -> Self {
        Self {
            ptr: pcwstr.as_ptr() as *mut u16,
        }
    }

    /// Converts the wide string to a Rust string slice
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_string(&self) -> Option<String> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        let os_string = OsString::from_wide(slice);
        Some(os_string.to_string_lossy().into_owned())
    }

    /// Converts the wide string to an `OsString`
    ///
    /// # Safety
    ///
    /// The caller must ensure the pointer is valid and points to a null-terminated wide string.
    pub unsafe fn to_os_string(&self) -> Option<OsString> {
        if self.ptr.is_null() {
            return None;
        }

        let mut len = 0;
        while unsafe { *self.ptr.add(len) } != 0 {
            len += 1;
        }

        let slice = unsafe { std::slice::from_raw_parts(self.ptr, len) };
        Some(OsString::from_wide(slice))
    }

    /// Returns the raw pointer without transferring ownership
    pub fn as_ptr(&self) -> *mut u16 {
        self.ptr
    }

    /// Returns the raw pointer and transfers ownership to the caller
    pub fn into_raw(mut self) -> *mut u16 {
        let ptr = self.ptr;
        self.ptr = ptr::null_mut();
        ptr
    }

    /// Checks if the pointer is null
    pub fn is_null(&self) -> bool {
        self.ptr.is_null()
    }

    /// Converts to a `PCWSTR` for use with Windows APIs
    pub fn as_pcwstr(&self) -> PCWSTR {
        PCWSTR(self.ptr)
    }
}

impl Drop for CalleeAllocatedWString {
    fn drop(&mut self) {
        if !self.ptr.is_null() {
            unsafe {
                CoTaskMemFree(Some(self.ptr.cast()));
            }
            self.ptr = ptr::null_mut();
        }
    }
}

impl Default for CalleeAllocatedWString {
    fn default() -> Self {
        Self {
            ptr: ptr::null_mut(),
        }
    }
}

impl Clone for CalleeAllocatedWString {
    /// Creates a shallow copy of the pointer.
    ///
    /// # Safety
    ///
    /// This creates two instances that will both attempt to free the same memory
    /// when dropped, potentially causing double-free errors. Use with extreme caution.
    fn clone(&self) -> Self {
        Self { ptr: self.ptr }
    }
}
