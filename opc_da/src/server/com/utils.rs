pub struct PointerReader;

pub struct PointerWriter;

pub trait TryWritePointer<T> {
    type Error;

    fn try_write(value: T, pointer: *mut T) -> Result<(), Self::Error>;
}

pub trait TryWriteInto<T, R> {
    type Error;

    fn try_write_into(value: T, pointer: R) -> Result<(), Self::Error>;
}

pub trait TryWriteTo<T, R> {
    type Error;

    fn try_write_to(value: T) -> Result<R, Self::Error>;
}

pub trait TryWriteArray<T, R = T> {
    type Error;

    fn try_write_array(values: &[T], pointer: *mut R) -> Result<(), Self::Error>;
}

pub trait TryWriteArrayPointer<T, R = T> {
    type Error;

    fn try_write_array_pointer(values: &[T], pointer: *mut *mut R) -> Result<(), Self::Error>;
}

pub trait TryRead<T, R = T> {
    type Error;

    fn try_read(pointer: *const T) -> Result<R, Self::Error>;
}

pub trait TryReadArray<T, R = T> {
    type Error;

    fn try_read_array(count: u32, pointer: *const T) -> Result<Vec<R>, Self::Error>;
}

impl<T: Sized> TryRead<T> for PointerReader {
    type Error = windows::core::Error;

    fn try_read(pointer: *const T) -> Result<T, Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        Ok(unsafe { pointer.read() })
    }
}

impl<T: Sized> TryWritePointer<T> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write(value: T, pointer: *mut T) -> Result<(), Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
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

impl<T: Sized + windows::core::Interface> TryWriteInto<T, windows::core::OutRef<'_, T>>
    for PointerWriter
{
    type Error = windows::core::Error;

    fn try_write_into(value: T, pointer: windows::core::OutRef<'_, T>) -> Result<(), Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        pointer.write(Some(value))?;

        Ok(())
    }
}

/// Allocates memory for a string and writes it to the provided pointer.  
///
/// # Safety  
/// The caller is responsible for freeing the allocated memory using `CoTaskMemFree`.  
impl<T: AsRef<str>> TryWriteInto<T, *mut windows::core::PWSTR> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write_into(value: T, pointer: *mut windows::core::PWSTR) -> Result<(), Self::Error> {
        let p = value
            .as_ref()
            .encode_utf16()
            .chain(core::iter::once(0))
            .collect::<Vec<u16>>();

        let ptr = unsafe {
            windows::Win32::System::Com::CoTaskMemAlloc(p.len() * core::mem::size_of::<u16>())
        };

        if ptr.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the string",
            ));
        }

        unsafe {
            core::ptr::copy_nonoverlapping(p.as_ptr(), ptr as *mut u16, p.len());
            *pointer = windows::core::PWSTR(ptr as *mut u16);
        }

        Ok(())
    }
}

impl<'a, T: AsRef<[&'a str]>> TryWriteInto<T, *mut *mut windows::core::PWSTR> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write_into(
        value: T,
        pointer: *mut *mut windows::core::PWSTR,
    ) -> Result<(), Self::Error> {
        let mut strings = Vec::with_capacity(value.as_ref().len());
        for s in value.as_ref() {
            let p = s
                .encode_utf16()
                .chain(core::iter::once(0))
                .collect::<Vec<u16>>();
            let ptr = unsafe {
                windows::Win32::System::Com::CoTaskMemAlloc(p.len() * core::mem::size_of::<u16>())
            };

            if ptr.is_null() {
                return Err(windows::core::Error::new(
                    windows::Win32::Foundation::E_OUTOFMEMORY,
                    "Failed to allocate memory for the string",
                ));
            }

            unsafe {
                core::ptr::copy_nonoverlapping(p.as_ptr(), ptr as *mut u16, p.len());
                strings.push(windows::core::PWSTR(ptr as *mut u16));
            }
        }

        let ptr = unsafe {
            windows::Win32::System::Com::CoTaskMemAlloc(
                strings.len() * core::mem::size_of::<windows::core::PWSTR>(),
            )
        };

        if ptr.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the array of strings",
            ));
        }

        unsafe {
            core::ptr::copy_nonoverlapping(
                strings.as_ptr(),
                ptr as *mut windows::core::PWSTR,
                strings.len(),
            );
            *pointer = ptr as _;
        }

        Ok(())
    }
}

impl<T, W: TryWritePointer<T, Error = windows::core::Error>> TryWriteTo<T, *mut T> for W {
    type Error = windows::core::Error;

    fn try_write_to(value: T) -> windows::core::Result<*mut T> {
        let ptr: *mut T = core::ptr::null_mut();
        Self::try_write(value, ptr)?;
        Ok(ptr)
    }
}

impl<T: AsRef<str>> TryWriteTo<T, windows::core::PWSTR> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write_to(value: T) -> windows::core::Result<windows::core::PWSTR> {
        let ptr: *mut windows::core::PWSTR = core::ptr::null_mut();
        Self::try_write_into(value, ptr)?;
        Ok(unsafe { *ptr })
    }
}

impl<T> TryReadArray<T> for PointerReader {
    type Error = windows::core::Error;

    fn try_read_array(count: u32, pointer: *const T) -> Result<Vec<T>, Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                result.push(pointer.add(i as usize).read());
            }
        }
        Ok(result)
    }
}

impl<T> TryWriteArray<T> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write_array(values: &[T], pointer: *mut T) -> Result<(), Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        unsafe {
            core::ptr::copy_nonoverlapping(values.as_ptr(), pointer, values.len());
        }

        Ok(())
    }
}

impl<T> TryWriteArrayPointer<T> for PointerWriter {
    type Error = windows::core::Error;

    fn try_write_array_pointer(values: &[T], pointer: *mut *mut T) -> Result<(), Self::Error> {
        if pointer.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Null pointer passed for 'pointer'",
            ));
        }

        let size = core::mem::size_of_val(values);
        let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };

        if ptr.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "Failed to allocate memory for the array",
            ));
        }

        unsafe {
            core::ptr::copy_nonoverlapping(values.as_ptr(), ptr as *mut T, values.len());
            *pointer = ptr as *mut T;
        }

        Ok(())
    }
}

impl TryReadArray<windows::core::PWSTR, String> for PointerReader {
    type Error = windows::core::Error;

    fn try_read_array(
        count: u32,
        pointer: *const windows::core::PWSTR,
    ) -> Result<Vec<String>, Self::Error> {
        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                let pwstr = pointer.add(i as usize).read();
                if pwstr.is_null() {
                    return Err(windows::core::Error::new(
                        windows::Win32::Foundation::E_POINTER,
                        "Null pointer encountered while reading string",
                    ));
                }
                result.push(pwstr.to_string()?);
            }
        }

        Ok(result)
    }
}

impl TryReadArray<windows::core::PCWSTR, String> for PointerReader {
    type Error = windows::core::Error;

    fn try_read_array(
        count: u32,
        pointer: *const windows::core::PCWSTR,
    ) -> Result<Vec<String>, Self::Error> {
        let mut result = Vec::with_capacity(count as usize);
        unsafe {
            for i in 0..count {
                let pwstr = pointer.add(i as usize).read();
                if pwstr.is_null() {
                    return Err(windows::core::Error::new(
                        windows::Win32::Foundation::E_POINTER,
                        "Null pointer encountered while reading string",
                    ));
                }
                result.push(pwstr.to_string()?);
            }
        }

        Ok(result)
    }
}
