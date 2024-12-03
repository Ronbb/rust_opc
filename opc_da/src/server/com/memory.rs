use windows::Win32::{Foundation::E_POINTER, System::Com::CoTaskMemAlloc};

pub trait IntoRef<Ref> {
    fn into_ref(self) -> windows_core::Result<Ref>;
}

pub trait IntoArrayRef<Ref> {
    fn into_array_ref(self, count: u32) -> windows_core::Result<Ref>;
}

pub trait IntoComArrayRef<Ref> {
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<Ref>;
}

pub trait FreeRaw {
    fn free_raw(self);
}

#[macro_export]
macro_rules! safe_call {
    ($call:expr, $($p:ident),*) => {
        let result = $call;

        if result.is_err() {
            $($p.free_raw();)*
        }

        result
    };
}

impl<'a, P> IntoRef<&'a P> for *const P {
    #[inline(always)]
    fn into_ref(self) -> windows_core::Result<&'a P> {
        if self.is_null() {
            Err(windows_core::Error::from_hresult(E_POINTER))
        } else {
            Ok(unsafe { &*self })
        }
    }
}

impl<'a, P> IntoRef<&'a mut P> for *mut P {
    #[inline(always)]
    fn into_ref(self) -> windows_core::Result<&'a mut P> {
        if self.is_null() {
            Err(windows_core::Error::from_hresult(E_POINTER))
        } else {
            Ok(unsafe { &mut *self })
        }
    }
}

impl<'a, P> IntoArrayRef<&'a mut [P]> for *mut P {
    #[inline(always)]
    fn into_array_ref(self, count: u32) -> windows_core::Result<&'a mut [P]> {
        if self.is_null() {
            Err(windows_core::Error::from_hresult(E_POINTER))
        } else {
            unsafe { Ok(std::slice::from_raw_parts_mut(self, count as usize)) }
        }
    }
}

impl<P> FreeRaw for *mut *mut P {
    #[inline(always)]
    fn free_raw(self) {
        unsafe {
            windows::Win32::System::Com::CoTaskMemFree(if self.is_null() {
                None
            } else {
                Some(*self as _)
            });
        }
    }
}

impl<'a, P> IntoComArrayRef<&'a [P]> for *const P {
    #[inline(always)]
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<&'a [P]> {
        if self.is_null() {
            Err(windows_core::Error::from_hresult(E_POINTER))
        } else {
            Ok(unsafe { core::slice::from_raw_parts(self, count as usize) })
        }
    }
}

impl<'a, P> IntoComArrayRef<&'a mut [P]> for *mut *mut P {
    #[inline(always)]
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<&'a mut [P]> {
        if self.is_null() {
            Err(windows_core::Error::from_hresult(E_POINTER))
        } else {
            unsafe {
                let new_pointer =
                    CoTaskMemAlloc(std::mem::size_of::<P>() * count as usize) as *mut P;

                if new_pointer.is_null() {
                    return Err(windows_core::Error::from_hresult(E_POINTER));
                } else {
                    *self = new_pointer;
                }

                Ok(std::slice::from_raw_parts_mut(new_pointer, count as usize))
            }
        }
    }
}

impl<'a, P1, P2> IntoComArrayRef<Vec<(&'a P1, &'a P2)>> for (*const P1, *const P2) {
    #[inline(always)]
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<Vec<(&'a P1, &'a P2)>> {
        let (p0, p1) = self;

        Ok(p0
            .into_com_array_ref(count)?
            .iter()
            .zip(p1.into_com_array_ref(count)?.iter())
            .collect())
    }
}

impl<'a, P1, P2> IntoComArrayRef<Vec<(&'a P1, &'a mut P2)>> for (*const P1, *mut *mut P2) {
    #[inline(always)]
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<Vec<(&'a P1, &'a mut P2)>> {
        let (p0, p1) = self;

        Ok(p0
            .into_com_array_ref(count)?
            .iter()
            .zip(p1.into_com_array_ref(count)?.iter_mut())
            .collect())
    }
}

impl<'a, P1, P2> IntoComArrayRef<Vec<(&'a mut P1, &'a mut P2)>> for (*mut *mut P1, *mut *mut P2) {
    #[inline(always)]
    fn into_com_array_ref(self, count: u32) -> windows_core::Result<Vec<(&'a mut P1, &'a mut P2)>> {
        let (p0, p1) = self;

        Ok(p0
            .into_com_array_ref(count)?
            .iter_mut()
            .zip(p1.into_com_array_ref(count)?.iter_mut())
            .collect())
    }
}

impl<'a, C1, M1, M2> IntoComArrayRef<Vec<(&'a C1, (&'a mut M1, &'a mut M2))>>
    for (*const C1, (*mut *mut M1, *mut *mut M2))
{
    #[inline(always)]
    fn into_com_array_ref(
        self,
        count: u32,
    ) -> windows_core::Result<Vec<(&'a C1, (&'a mut M1, &'a mut M2))>> {
        let (c, m) = self;

        Ok(c.into_com_array_ref(count)?
            .iter()
            .zip(m.into_com_array_ref(count)?)
            .collect())
    }
}

impl<'a, C1, C2, M1> IntoComArrayRef<Vec<((&'a C1, &'a C2), &'a mut M1)>>
    for ((*const C1, *const C2), *mut *mut M1)
{
    #[inline(always)]
    fn into_com_array_ref(
        self,
        count: u32,
    ) -> windows_core::Result<Vec<((&'a C1, &'a C2), &'a mut M1)>> {
        let (c, m) = self;

        Ok(c.into_com_array_ref(count)?
            .into_iter()
            .zip(m.into_com_array_ref(count)?)
            .collect())
    }
}

impl<'a, C1, C2, M1, M2> IntoComArrayRef<Vec<((&'a C1, &'a C2), (&'a mut M1, &'a mut M2))>>
    for ((*const C1, *const C2), (*mut *mut M1, *mut *mut M2))
{
    #[inline(always)]
    fn into_com_array_ref(
        self,
        count: u32,
    ) -> windows_core::Result<Vec<((&'a C1, &'a C2), (&'a mut M1, &'a mut M2))>> {
        let (c, m) = self;

        Ok(c.into_com_array_ref(count)?
            .into_iter()
            .zip(m.into_com_array_ref(count)?)
            .collect())
    }
}
