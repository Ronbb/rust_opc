use std::mem::size_of;

use windows::Win32::System::Com::CoTaskMemAlloc;
use windows_core::PWSTR;

pub fn com_alloc_str(s: &str) -> PWSTR {
    let v: Vec<u16> = s.encode_utf16().chain(Some(0)).collect();

    PWSTR::from_raw(com_alloc_v(&v))
}

pub fn com_alloc_v<T>(v: &Vec<T>) -> *mut T {
    let size = v.len() * size_of::<T>();

    unsafe {
        let ptr = CoTaskMemAlloc(size) as *mut T;
        if ptr.is_null() {
            panic!("CoTaskMemAlloc failed");
        }

        std::ptr::copy_nonoverlapping(v.as_ptr(), ptr, v.len());

        ptr
    }
}
