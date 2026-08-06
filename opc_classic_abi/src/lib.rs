//! Private shared ABI definitions used to compose OPC Classic COM identities.
//!
//! This crate is not published and must never appear in a safe client or server
//! signature. Public crates translate these definitions to `opc_classic_types`.

#![allow(non_camel_case_types, non_snake_case)]

const E_POINTER: windows_core::HRESULT = windows_core::HRESULT(0x8000_4003_u32 as i32);

// This mirrors the narrow interface shape emitted by windows-bindgen. The
// generated output itself uses these windows-core macros; keeping the single
// interface here lets us harden its out-parameter thunks without exposing the
// full generated OPC Common ABI to the domain crates.
windows_core::imp::define_interface!(
    IOPCCommon,
    IOPCCommon_Vtbl,
    0xf31dfde2_07b6_11d2_b2d8_0060083ba1fb
);
windows_core::imp::interface_hierarchy!(IOPCCommon, windows_core::IUnknown);

#[repr(C)]
#[doc(hidden)]
pub struct IOPCCommon_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub SetLocaleID:
        unsafe extern "system" fn(*mut core::ffi::c_void, u32) -> windows_core::HRESULT,
    pub GetLocaleID:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub QueryAvailableLocaleIDs: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut u32,
        *mut *mut u32,
    ) -> windows_core::HRESULT,
    pub GetErrorString: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::HRESULT,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub SetClientName: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
}

#[doc(hidden)]
pub trait IOPCCommon_Impl: windows_core::IUnknownImpl {
    fn SetLocaleID(&self, locale: u32) -> windows_core::Result<()>;
    fn GetLocaleID(&self) -> windows_core::Result<u32>;
    fn QueryAvailableLocaleIDs(
        &self,
        count: *mut u32,
        locales: *mut *mut u32,
    ) -> windows_core::Result<()>;
    fn GetErrorString(
        &self,
        error: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR>;
    fn SetClientName(&self, name: &windows_core::PCWSTR) -> windows_core::Result<()>;
}

impl IOPCCommon_Vtbl {
    pub const fn new<Identity: IOPCCommon_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn set_locale<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            locale: u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this = &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::SetLocaleID(this, locale).into()
            }
        }

        unsafe extern "system" fn get_locale<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            locale: *mut u32,
        ) -> windows_core::HRESULT {
            if locale.is_null() {
                return E_POINTER;
            }

            unsafe {
                locale.write(0);
                let this = &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCCommon_Impl::GetLocaleID(this) {
                    Ok(value) => {
                        locale.write(value);
                        windows_core::HRESULT(0)
                    }
                    Err(error) => error.into(),
                }
            }
        }

        unsafe extern "system" fn available_locales<
            Identity: IOPCCommon_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            count: *mut u32,
            locales: *mut *mut u32,
        ) -> windows_core::HRESULT {
            if count.is_null() || locales.is_null() {
                return E_POINTER;
            }

            unsafe {
                count.write(0);
                locales.write(core::ptr::null_mut());
                let this = &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::QueryAvailableLocaleIDs(this, count, locales).into()
            }
        }

        unsafe extern "system" fn error_string<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            error: windows_core::HRESULT,
            output: *mut windows_core::PWSTR,
        ) -> windows_core::HRESULT {
            if output.is_null() {
                return E_POINTER;
            }

            unsafe {
                output.write(windows_core::PWSTR::null());
                let this = &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCCommon_Impl::GetErrorString(this, error) {
                    Ok(value) => {
                        output.write(value);
                        windows_core::HRESULT(0)
                    }
                    Err(error) => error.into(),
                }
            }
        }

        unsafe extern "system" fn set_client_name<
            Identity: IOPCCommon_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            name: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this = &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::SetClientName(this, &name).into()
            }
        }

        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            SetLocaleID: set_locale::<Identity, OFFSET>,
            GetLocaleID: get_locale::<Identity, OFFSET>,
            QueryAvailableLocaleIDs: available_locales::<Identity, OFFSET>,
            GetErrorString: error_string::<Identity, OFFSET>,
            SetClientName: set_client_name::<Identity, OFFSET>,
        }
    }

    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCCommon as windows_core::Interface>::IID
    }
}

impl windows_core::RuntimeName for IOPCCommon {}
