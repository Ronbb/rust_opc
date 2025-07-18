// Bindings generated by `windows-bindgen` 0.62.1

#![allow(
    non_snake_case,
    non_upper_case_globals,
    non_camel_case_types,
    dead_code,
    clippy::all
)]

windows_core::imp::define_interface!(
    IOPCCommon,
    IOPCCommon_Vtbl,
    0xf31dfde2_07b6_11d2_b2d8_0060083ba1fb
);
windows_core::imp::interface_hierarchy!(IOPCCommon, windows_core::IUnknown);
impl IOPCCommon {
    pub unsafe fn SetLocaleID(&self, dwlcid: u32) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).SetLocaleID)(
                windows_core::Interface::as_raw(self),
                dwlcid,
            )
            .ok()
        }
    }
    pub unsafe fn GetLocaleID(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetLocaleID)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn QueryAvailableLocaleIDs(
        &self,
        pdwcount: *mut u32,
        pdwlcid: *mut *mut u32,
    ) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).QueryAvailableLocaleIDs)(
                windows_core::Interface::as_raw(self),
                pdwcount as _,
                pdwlcid as _,
            )
            .ok()
        }
    }
    pub unsafe fn GetErrorString(
        &self,
        dwerror: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetErrorString)(
                windows_core::Interface::as_raw(self),
                dwerror,
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn SetClientName<P0>(&self, szname: P0) -> windows_core::Result<()>
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).SetClientName)(
                windows_core::Interface::as_raw(self),
                szname.param().abi(),
            )
            .ok()
        }
    }
}
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
pub trait IOPCCommon_Impl: windows_core::IUnknownImpl {
    fn SetLocaleID(&self, dwlcid: u32) -> windows_core::Result<()>;
    fn GetLocaleID(&self) -> windows_core::Result<u32>;
    fn QueryAvailableLocaleIDs(
        &self,
        pdwcount: *mut u32,
        pdwlcid: *mut *mut u32,
    ) -> windows_core::Result<()>;
    fn GetErrorString(
        &self,
        dwerror: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR>;
    fn SetClientName(&self, szname: &windows_core::PCWSTR) -> windows_core::Result<()>;
}
impl IOPCCommon_Vtbl {
    pub const fn new<Identity: IOPCCommon_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn SetLocaleID<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            dwlcid: u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::SetLocaleID(this, core::mem::transmute_copy(&dwlcid)).into()
            }
        }
        unsafe extern "system" fn GetLocaleID<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            pdwlcid: *mut u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCCommon_Impl::GetLocaleID(this) {
                    Ok(ok__) => {
                        pdwlcid.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        unsafe extern "system" fn QueryAvailableLocaleIDs<
            Identity: IOPCCommon_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            pdwcount: *mut u32,
            pdwlcid: *mut *mut u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::QueryAvailableLocaleIDs(
                    this,
                    core::mem::transmute_copy(&pdwcount),
                    core::mem::transmute_copy(&pdwlcid),
                )
                .into()
            }
        }
        unsafe extern "system" fn GetErrorString<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            dwerror: windows_core::HRESULT,
            ppstring: *mut windows_core::PWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCCommon_Impl::GetErrorString(this, core::mem::transmute_copy(&dwerror)) {
                    Ok(ok__) => {
                        ppstring.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        unsafe extern "system" fn SetClientName<Identity: IOPCCommon_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            szname: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCCommon_Impl::SetClientName(this, core::mem::transmute(&szname)).into()
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            SetLocaleID: SetLocaleID::<Identity, OFFSET>,
            GetLocaleID: GetLocaleID::<Identity, OFFSET>,
            QueryAvailableLocaleIDs: QueryAvailableLocaleIDs::<Identity, OFFSET>,
            GetErrorString: GetErrorString::<Identity, OFFSET>,
            SetClientName: SetClientName::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCCommon as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IOPCCommon {}
windows_core::imp::define_interface!(
    IOPCEnumGUID,
    IOPCEnumGUID_Vtbl,
    0x55c382c8_21c7_4e88_96c1_becfb1e3f483
);
windows_core::imp::interface_hierarchy!(IOPCEnumGUID, windows_core::IUnknown);
impl IOPCEnumGUID {
    pub unsafe fn Next(
        &self,
        rgelt: &mut [windows_core::GUID],
        pceltfetched: *mut u32,
    ) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).Next)(
                windows_core::Interface::as_raw(self),
                rgelt.len().try_into().unwrap(),
                core::mem::transmute(rgelt.as_ptr()),
                pceltfetched as _,
            )
            .ok()
        }
    }
    pub unsafe fn Skip(&self, celt: u32) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).Skip)(
                windows_core::Interface::as_raw(self),
                celt,
            )
            .ok()
        }
    }
    pub unsafe fn Reset(&self) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).Reset)(windows_core::Interface::as_raw(self))
                .ok()
        }
    }
    pub unsafe fn Clone(&self) -> windows_core::Result<IOPCEnumGUID> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).Clone)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }
}
#[repr(C)]
#[doc(hidden)]
pub struct IOPCEnumGUID_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Next: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *mut windows_core::GUID,
        *mut u32,
    ) -> windows_core::HRESULT,
    pub Skip: unsafe extern "system" fn(*mut core::ffi::c_void, u32) -> windows_core::HRESULT,
    pub Reset: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub Clone: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
pub trait IOPCEnumGUID_Impl: windows_core::IUnknownImpl {
    fn Next(
        &self,
        celt: u32,
        rgelt: *mut windows_core::GUID,
        pceltfetched: *mut u32,
    ) -> windows_core::Result<()>;
    fn Skip(&self, celt: u32) -> windows_core::Result<()>;
    fn Reset(&self) -> windows_core::Result<()>;
    fn Clone(&self) -> windows_core::Result<IOPCEnumGUID>;
}
impl IOPCEnumGUID_Vtbl {
    pub const fn new<Identity: IOPCEnumGUID_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn Next<Identity: IOPCEnumGUID_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            celt: u32,
            rgelt: *mut windows_core::GUID,
            pceltfetched: *mut u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCEnumGUID_Impl::Next(
                    this,
                    core::mem::transmute_copy(&celt),
                    core::mem::transmute_copy(&rgelt),
                    core::mem::transmute_copy(&pceltfetched),
                )
                .into()
            }
        }
        unsafe extern "system" fn Skip<Identity: IOPCEnumGUID_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            celt: u32,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCEnumGUID_Impl::Skip(this, core::mem::transmute_copy(&celt)).into()
            }
        }
        unsafe extern "system" fn Reset<Identity: IOPCEnumGUID_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCEnumGUID_Impl::Reset(this).into()
            }
        }
        unsafe extern "system" fn Clone<Identity: IOPCEnumGUID_Impl, const OFFSET: isize>(
            this: *mut core::ffi::c_void,
            ppenum: *mut *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCEnumGUID_Impl::Clone(this) {
                    Ok(ok__) => {
                        ppenum.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            Next: Next::<Identity, OFFSET>,
            Skip: Skip::<Identity, OFFSET>,
            Reset: Reset::<Identity, OFFSET>,
            Clone: Clone::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCEnumGUID as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IOPCEnumGUID {}
windows_core::imp::define_interface!(
    IOPCServerList,
    IOPCServerList_Vtbl,
    0x13486d50_4821_11d2_a494_3cb306c10000
);
windows_core::imp::interface_hierarchy!(IOPCServerList, windows_core::IUnknown);
impl IOPCServerList {
    pub unsafe fn EnumClassesOfCategories(
        &self,
        rgcatidimpl: &[windows_core::GUID],
        rgcatidreq: &[windows_core::GUID],
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumGUID> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).EnumClassesOfCategories)(
                windows_core::Interface::as_raw(self),
                rgcatidimpl.len().try_into().unwrap(),
                core::mem::transmute(rgcatidimpl.as_ptr()),
                rgcatidreq.len().try_into().unwrap(),
                core::mem::transmute(rgcatidreq.as_ptr()),
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).GetClassDetails)(
                windows_core::Interface::as_raw(self),
                clsid,
                ppszprogid as _,
                ppszusertype as _,
            )
            .ok()
        }
    }
    pub unsafe fn CLSIDFromProgID<P0>(
        &self,
        szprogid: P0,
    ) -> windows_core::Result<windows_core::GUID>
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).CLSIDFromProgID)(
                windows_core::Interface::as_raw(self),
                szprogid.param().abi(),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
#[doc(hidden)]
pub struct IOPCServerList_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub EnumClassesOfCategories: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *const windows_core::GUID,
        u32,
        *const windows_core::GUID,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetClassDetails: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        *mut windows_core::PWSTR,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub CLSIDFromProgID: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *mut windows_core::GUID,
    ) -> windows_core::HRESULT,
}
pub trait IOPCServerList_Impl: windows_core::IUnknownImpl {
    fn EnumClassesOfCategories(
        &self,
        cimplemented: u32,
        rgcatidimpl: *const windows_core::GUID,
        crequired: u32,
        rgcatidreq: *const windows_core::GUID,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumGUID>;
    fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()>;
    fn CLSIDFromProgID(
        &self,
        szprogid: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::GUID>;
}
impl IOPCServerList_Vtbl {
    pub const fn new<Identity: IOPCServerList_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn EnumClassesOfCategories<
            Identity: IOPCServerList_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            cimplemented: u32,
            rgcatidimpl: *const windows_core::GUID,
            crequired: u32,
            rgcatidreq: *const windows_core::GUID,
            ppenumclsid: *mut *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCServerList_Impl::EnumClassesOfCategories(
                    this,
                    core::mem::transmute_copy(&cimplemented),
                    core::mem::transmute_copy(&rgcatidimpl),
                    core::mem::transmute_copy(&crequired),
                    core::mem::transmute_copy(&rgcatidreq),
                ) {
                    Ok(ok__) => {
                        ppenumclsid.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        unsafe extern "system" fn GetClassDetails<
            Identity: IOPCServerList_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            clsid: *const windows_core::GUID,
            ppszprogid: *mut windows_core::PWSTR,
            ppszusertype: *mut windows_core::PWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCServerList_Impl::GetClassDetails(
                    this,
                    core::mem::transmute_copy(&clsid),
                    core::mem::transmute_copy(&ppszprogid),
                    core::mem::transmute_copy(&ppszusertype),
                )
                .into()
            }
        }
        unsafe extern "system" fn CLSIDFromProgID<
            Identity: IOPCServerList_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            szprogid: windows_core::PCWSTR,
            clsid: *mut windows_core::GUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCServerList_Impl::CLSIDFromProgID(this, core::mem::transmute(&szprogid)) {
                    Ok(ok__) => {
                        clsid.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            EnumClassesOfCategories: EnumClassesOfCategories::<Identity, OFFSET>,
            GetClassDetails: GetClassDetails::<Identity, OFFSET>,
            CLSIDFromProgID: CLSIDFromProgID::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCServerList as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IOPCServerList {}
windows_core::imp::define_interface!(
    IOPCServerList2,
    IOPCServerList2_Vtbl,
    0x9dd0b56c_ad9e_43ee_8305_487f3188bf7a
);
windows_core::imp::interface_hierarchy!(IOPCServerList2, windows_core::IUnknown);
impl IOPCServerList2 {
    pub unsafe fn EnumClassesOfCategories(
        &self,
        rgcatidimpl: &[windows_core::GUID],
        rgcatidreq: &[windows_core::GUID],
    ) -> windows_core::Result<IOPCEnumGUID> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).EnumClassesOfCategories)(
                windows_core::Interface::as_raw(self),
                rgcatidimpl.len().try_into().unwrap(),
                core::mem::transmute(rgcatidimpl.as_ptr()),
                rgcatidreq.len().try_into().unwrap(),
                core::mem::transmute(rgcatidreq.as_ptr()),
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
        ppszverindprogid: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()> {
        unsafe {
            (windows_core::Interface::vtable(self).GetClassDetails)(
                windows_core::Interface::as_raw(self),
                clsid,
                ppszprogid as _,
                ppszusertype as _,
                ppszverindprogid as _,
            )
            .ok()
        }
    }
    pub unsafe fn CLSIDFromProgID<P0>(
        &self,
        szprogid: P0,
    ) -> windows_core::Result<windows_core::GUID>
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).CLSIDFromProgID)(
                windows_core::Interface::as_raw(self),
                szprogid.param().abi(),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
#[doc(hidden)]
pub struct IOPCServerList2_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub EnumClassesOfCategories: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *const windows_core::GUID,
        u32,
        *const windows_core::GUID,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetClassDetails: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        *mut windows_core::PWSTR,
        *mut windows_core::PWSTR,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub CLSIDFromProgID: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *mut windows_core::GUID,
    ) -> windows_core::HRESULT,
}
pub trait IOPCServerList2_Impl: windows_core::IUnknownImpl {
    fn EnumClassesOfCategories(
        &self,
        cimplemented: u32,
        rgcatidimpl: *const windows_core::GUID,
        crequired: u32,
        rgcatidreq: *const windows_core::GUID,
    ) -> windows_core::Result<IOPCEnumGUID>;
    fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
        ppszverindprogid: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()>;
    fn CLSIDFromProgID(
        &self,
        szprogid: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::GUID>;
}
impl IOPCServerList2_Vtbl {
    pub const fn new<Identity: IOPCServerList2_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn EnumClassesOfCategories<
            Identity: IOPCServerList2_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            cimplemented: u32,
            rgcatidimpl: *const windows_core::GUID,
            crequired: u32,
            rgcatidreq: *const windows_core::GUID,
            ppenumclsid: *mut *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCServerList2_Impl::EnumClassesOfCategories(
                    this,
                    core::mem::transmute_copy(&cimplemented),
                    core::mem::transmute_copy(&rgcatidimpl),
                    core::mem::transmute_copy(&crequired),
                    core::mem::transmute_copy(&rgcatidreq),
                ) {
                    Ok(ok__) => {
                        ppenumclsid.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        unsafe extern "system" fn GetClassDetails<
            Identity: IOPCServerList2_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            clsid: *const windows_core::GUID,
            ppszprogid: *mut windows_core::PWSTR,
            ppszusertype: *mut windows_core::PWSTR,
            ppszverindprogid: *mut windows_core::PWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCServerList2_Impl::GetClassDetails(
                    this,
                    core::mem::transmute_copy(&clsid),
                    core::mem::transmute_copy(&ppszprogid),
                    core::mem::transmute_copy(&ppszusertype),
                    core::mem::transmute_copy(&ppszverindprogid),
                )
                .into()
            }
        }
        unsafe extern "system" fn CLSIDFromProgID<
            Identity: IOPCServerList2_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            szprogid: windows_core::PCWSTR,
            clsid: *mut windows_core::GUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                match IOPCServerList2_Impl::CLSIDFromProgID(this, core::mem::transmute(&szprogid)) {
                    Ok(ok__) => {
                        clsid.write(core::mem::transmute(ok__));
                        windows_core::HRESULT(0)
                    }
                    Err(err) => err.into(),
                }
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            EnumClassesOfCategories: EnumClassesOfCategories::<Identity, OFFSET>,
            GetClassDetails: GetClassDetails::<Identity, OFFSET>,
            CLSIDFromProgID: CLSIDFromProgID::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCServerList2 as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IOPCServerList2 {}
windows_core::imp::define_interface!(
    IOPCShutdown,
    IOPCShutdown_Vtbl,
    0xf31dfde1_07b6_11d2_b2d8_0060083ba1fb
);
windows_core::imp::interface_hierarchy!(IOPCShutdown, windows_core::IUnknown);
impl IOPCShutdown {
    pub unsafe fn ShutdownRequest<P0>(&self, szreason: P0) -> windows_core::Result<()>
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).ShutdownRequest)(
                windows_core::Interface::as_raw(self),
                szreason.param().abi(),
            )
            .ok()
        }
    }
}
#[repr(C)]
#[doc(hidden)]
pub struct IOPCShutdown_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub ShutdownRequest: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
}
pub trait IOPCShutdown_Impl: windows_core::IUnknownImpl {
    fn ShutdownRequest(&self, szreason: &windows_core::PCWSTR) -> windows_core::Result<()>;
}
impl IOPCShutdown_Vtbl {
    pub const fn new<Identity: IOPCShutdown_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn ShutdownRequest<
            Identity: IOPCShutdown_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            szreason: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IOPCShutdown_Impl::ShutdownRequest(this, core::mem::transmute(&szreason)).into()
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            ShutdownRequest: ShutdownRequest::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IOPCShutdown as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IOPCShutdown {}
