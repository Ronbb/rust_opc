use windows::Win32::System::Com::CoTaskMemAlloc;
use windows_core::PWSTR;

pub fn copy_to_com_string(s: &str) -> windows_core::Result<PWSTR> {
    let v: Vec<u16> = s.encode_utf16().chain(Some(0)).collect();

    unsafe {
        let ptr = CoTaskMemAlloc(v.len() * std::mem::size_of::<u16>()) as *mut u16;
        if ptr.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory",
            ));
        }

        std::ptr::copy_nonoverlapping(v.as_ptr(), ptr, v.len());

        Ok(PWSTR(ptr))
    }
}

pub fn copy_to_com_pointer<T>(
    vec: &[T],
    count: *mut u32,
    pointer: *mut *mut T,
) -> windows_core::Result<()> {
    unsafe {
        if count.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'count'",
            ));
        }

        *count = vec.len() as u32;

        if pointer.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        let size = vec.len() * std::mem::size_of::<T>();
        let allocated_ptr = CoTaskMemAlloc(size) as *mut T;
        if allocated_ptr.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory",
            ));
        }

        std::ptr::copy_nonoverlapping(vec.as_ptr(), allocated_ptr, vec.len());

        *pointer = allocated_ptr;
    }

    Ok(())
}

pub fn copy_to_pointer<T>(
    vec: &[T],
    count: *mut u32,
    pointer: *mut *mut T,
) -> windows_core::Result<()> {
    unsafe {
        if !count.is_null() {
            *count = vec.len() as u32;
        } else {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'count'",
            ));
        }

        if !pointer.is_null() && !(*pointer).is_null() {
            std::ptr::copy_nonoverlapping(vec.as_ptr(), *pointer, vec.len());
        } else {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }
    }

    Ok(())
}
