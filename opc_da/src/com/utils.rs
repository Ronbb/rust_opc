pub struct PointerReader;

pub struct PointerWriter;

pub trait Read<T, R = T> {
    fn read(pointer: *const T) -> R;
}

pub trait Write<T> {
    fn write(value: T, pointer: *mut T) -> windows_core::Result<()>;
}

pub trait WriteInto<T, R> {
    fn write_into(value: T, pointer: *mut R) -> windows_core::Result<()>;
}

pub trait WriteTo<T, R> {
    fn write_to(value: T) -> windows_core::Result<R>;
}

pub trait ReadArray<T, R = T> {
    fn read_array(count: u32, pointer: *const T) -> Vec<R>;
}

pub trait WriteArray<T, R = T> {
    fn write_array(values: &[T], pointer: *mut R) -> windows_core::Result<()>;
}

pub trait WriteArrayPointer<T, R = T> {
    fn write_array_pointer(values: &[T], pointer: *mut *mut R) -> windows_core::Result<()>;
}

pub trait TryReadArray<T, R = T> {
    type Error;

    fn try_read_array(count: u32, pointer: *const T) -> Result<Vec<R>, Self::Error>;
}

pub trait TryWriteArray<T, R = T> {
    type Error;

    fn try_write_array(values: &[T], pointer: *mut R) -> Result<(), Self::Error>;
}

impl<T: Sized> Read<T> for PointerReader {
    fn read(pointer: *const T) -> T {
        unsafe { pointer.read() }
    }
}

impl<T: Sized> Write<T> for PointerWriter {
    fn write(value: T, pointer: *mut T) -> windows_core::Result<()> {
        if pointer.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        unsafe {
            pointer.write(value);
        }

        Ok(())
    }
}

impl<T: AsRef<str>> WriteInto<T, windows_core::PWSTR> for PointerWriter {
    fn write_into(value: T, pointer: *mut windows_core::PWSTR) -> windows_core::Result<()> {
        let p = value
            .as_ref()
            .encode_utf16()
            .chain(core::iter::once(0))
            .collect::<Vec<u16>>();

        let ptr = unsafe {
            windows::Win32::System::Com::CoTaskMemAlloc(p.len() * std::mem::size_of::<u16>())
        };

        if ptr.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the string",
            ));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(p.as_ptr(), ptr as *mut u16, p.len());
            *pointer = windows_core::PWSTR::from_raw(ptr as *mut u16);
        }

        Ok(())
    }
}

impl<'a, T: AsRef<[&'a str]>> WriteInto<T, *mut windows_core::PWSTR> for PointerWriter {
    fn write_into(value: T, pointer: *mut *mut windows_core::PWSTR) -> windows_core::Result<()> {
        let mut strings = Vec::with_capacity(value.as_ref().len());
        for s in value.as_ref() {
            let p = s
                .encode_utf16()
                .chain(core::iter::once(0))
                .collect::<Vec<u16>>();
            let ptr = unsafe {
                windows::Win32::System::Com::CoTaskMemAlloc(p.len() * std::mem::size_of::<u16>())
            };

            if ptr.is_null() {
                return Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_OUTOFMEMORY,
                    "Failed to allocate memory for the string",
                ));
            }

            unsafe {
                std::ptr::copy_nonoverlapping(p.as_ptr(), ptr as *mut u16, p.len());
                strings.push(windows_core::PWSTR::from_raw(ptr as *mut u16));
            }
        }

        let ptr = unsafe {
            windows::Win32::System::Com::CoTaskMemAlloc(strings.len() * std::mem::size_of::<u16>())
        };

        if ptr.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the array of strings",
            ));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(
                strings.as_ptr(),
                ptr as *mut windows_core::PWSTR,
                strings.len(),
            );
            *pointer = ptr as _;
        }

        Ok(())
    }
}

impl<T, R, W: WriteInto<T, R>> WriteTo<T, *mut R> for W {
    fn write_to(value: T) -> windows_core::Result<*mut R> {
        let ptr: *mut R = std::ptr::null_mut();
        Self::write_into(value, ptr)?;
        Ok(ptr)
    }
}

impl<T: AsRef<str>> WriteTo<T, windows_core::PWSTR> for PointerWriter {
    fn write_to(value: T) -> windows_core::Result<windows_core::PWSTR> {
        let ptr: *mut windows_core::PWSTR = std::ptr::null_mut();
        Self::write_into(value, ptr)?;
        Ok(unsafe { *ptr })
    }
}

impl<T> ReadArray<T> for PointerReader {
    fn read_array(count: u32, pointer: *const T) -> Vec<T> {
        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                result.push(pointer.add(i as usize).read());
            }
        }
        result
    }
}

impl<T> WriteArray<T> for PointerWriter {
    fn write_array(values: &[T], pointer: *mut T) -> windows_core::Result<()> {
        if pointer.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(values.as_ptr(), pointer, values.len());
        }

        Ok(())
    }
}

impl<T> WriteArrayPointer<T> for PointerWriter {
    fn write_array_pointer(values: &[T], pointer: *mut *mut T) -> windows_core::Result<()> {
        if pointer.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        let size = std::mem::size_of::<T>() * values.len();
        let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };

        if ptr.is_null() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the array",
            ));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(values.as_ptr(), ptr as *mut T, values.len());
            *pointer = ptr as *mut T;
        }

        Ok(())
    }
}

impl TryReadArray<windows_core::PWSTR, String> for PointerReader {
    type Error = windows_core::Error;

    fn try_read_array(
        count: u32,
        pointer: *const windows_core::PWSTR,
    ) -> Result<Vec<String>, Self::Error> {
        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                result.push(pointer.add(i as usize).read().to_string()?);
            }
        }

        Ok(result)
    }
}

impl TryReadArray<windows_core::PCWSTR, String> for PointerReader {
    type Error = windows_core::Error;

    fn try_read_array(
        count: u32,
        pointer: *const windows_core::PCWSTR,
    ) -> Result<Vec<String>, Self::Error> {
        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                result.push(pointer.add(i as usize).read().to_string()?);
            }
        }

        Ok(result)
    }
}
