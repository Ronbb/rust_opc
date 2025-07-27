#[windows::core::implement(opc_comn_bindings::IOPCCommon)]
pub struct OPCCommon<T>(T)
where
    T: 'static + traits::OPCCommon;

impl<T: traits::OPCCommon> opc_comn_bindings::IOPCCommon_Impl for OPCCommon_Impl<T> {
    fn SetLocaleID(&self, dwlcid: u32) -> windows_core::Result<()> {
        self.0.set_locale_id(dwlcid)
    }

    fn GetLocaleID(&self) -> windows_core::Result<u32> {
        self.0.get_locale_id()
    }

    fn QueryAvailableLocaleIDs(
        &self,
        pdwcount: *mut u32,
        pdwlcid: *mut *mut u32,
    ) -> windows_core::Result<()> {
        let locale_ids = self.0.query_available_locale_ids()?;
        opc_classic_utils::write_caller_allocated_ptr!(pdwcount, locale_ids.len().try_into()?)?;
        opc_classic_utils::write_caller_allocated_array!(pdwlcid, locale_ids.as_slice())?;

        Ok(())
    }

    fn GetErrorString(
        &self,
        dwerror: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR> {
        opc_classic_utils::alloc_callee_wstring!(self.0.get_error_string(dwerror)?)
    }

    fn SetClientName(&self, szname: &windows_core::PCWSTR) -> windows_core::Result<()> {
        self.0.set_client_name(unsafe {
            opc_classic_utils::CallerAllocatedWString::from_pcwstr(*szname)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?)
    }
}

#[windows::core::implement(opc_comn_bindings::IOPCEnumGUID)]
pub struct OPCEnumGUID<T>(T)
where
    T: 'static + traits::OPCEnumGUID;

impl<T: traits::OPCEnumGUID> opc_comn_bindings::IOPCEnumGUID_Impl for OPCEnumGUID_Impl<T> {
    fn Next(
        &self,
        celt: u32,
        rgelt: *mut windows_core::GUID,
        pceltfetched: *mut u32,
    ) -> windows_core::Result<()> {
        let guids = self.0.next(celt)?;
        let fetched = guids.len() as u32;

        if !rgelt.is_null() {
            opc_classic_utils::copy_to_caller_array!(rgelt, &guids)?;
        }

        opc_classic_utils::write_caller_allocated_ptr!(pceltfetched, fetched)?;

        Ok(())
    }

    fn Skip(&self, celt: u32) -> windows_core::Result<()> {
        self.0.skip(celt)
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset()
    }

    fn Clone(&self) -> windows_core::Result<opc_comn_bindings::IOPCEnumGUID> {
        self.0.clone_enum()
    }
}

#[windows::core::implement(opc_comn_bindings::IOPCServerList)]
pub struct OPCServerList<T>(T)
where
    T: 'static + traits::OPCServerList;

impl<T: traits::OPCServerList> opc_comn_bindings::IOPCServerList_Impl for OPCServerList_Impl<T> {
    fn EnumClassesOfCategories(
        &self,
        cimplemented: u32,
        rgcatidimpl: *const windows_core::GUID,
        crequired: u32,
        rgcatidreq: *const windows_core::GUID,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumGUID> {
        let implemented = if !rgcatidimpl.is_null() {
            unsafe {
                opc_classic_utils::CallerAllocatedArray::new(
                    rgcatidimpl as *mut windows_core::GUID,
                    cimplemented as usize,
                )
                .as_slice()
                .unwrap_or(&[])
                .to_vec()
            }
        } else {
            Vec::new()
        };

        let required = if !rgcatidreq.is_null() {
            unsafe {
                opc_classic_utils::CallerAllocatedArray::new(
                    rgcatidreq as *mut windows_core::GUID,
                    crequired as usize,
                )
                .as_slice()
                .unwrap_or(&[])
                .to_vec()
            }
        } else {
            Vec::new()
        };

        self.0.enum_classes_of_categories(implemented, required)
    }

    fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()> {
        let (prog_id, user_type) = self.0.get_class_details(unsafe { *clsid })?;

        opc_classic_utils::write_caller_allocated_ptr!(
            ppszprogid,
            opc_classic_utils::alloc_callee_wstring!(prog_id)?
        )?;
        opc_classic_utils::write_caller_allocated_ptr!(
            ppszusertype,
            opc_classic_utils::alloc_callee_wstring!(user_type)?
        )?;

        Ok(())
    }

    fn CLSIDFromProgID(
        &self,
        szprogid: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::GUID> {
        let prog_id = unsafe {
            opc_classic_utils::CallerAllocatedWString::from_pcwstr(*szprogid)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.clsid_from_prog_id(prog_id)
    }
}

#[windows::core::implement(opc_comn_bindings::IOPCServerList2)]
pub struct OPCServerList2<T>(T)
where
    T: 'static + traits::OPCServerList2;

impl<T: traits::OPCServerList2> opc_comn_bindings::IOPCServerList2_Impl for OPCServerList2_Impl<T> {
    fn EnumClassesOfCategories(
        &self,
        cimplemented: u32,
        rgcatidimpl: *const windows_core::GUID,
        crequired: u32,
        rgcatidreq: *const windows_core::GUID,
    ) -> windows_core::Result<opc_comn_bindings::IOPCEnumGUID> {
        let implemented = if !rgcatidimpl.is_null() {
            unsafe {
                opc_classic_utils::CallerAllocatedArray::new(
                    rgcatidimpl as *mut windows_core::GUID,
                    cimplemented as usize,
                )
                .as_slice()
                .unwrap_or(&[])
                .to_vec()
            }
        } else {
            Vec::new()
        };

        let required = if !rgcatidreq.is_null() {
            unsafe {
                opc_classic_utils::CallerAllocatedArray::new(
                    rgcatidreq as *mut windows_core::GUID,
                    crequired as usize,
                )
                .as_slice()
                .unwrap_or(&[])
                .to_vec()
            }
        } else {
            Vec::new()
        };

        self.0.enum_classes_of_categories(implemented, required)
    }

    fn GetClassDetails(
        &self,
        clsid: *const windows_core::GUID,
        ppszprogid: *mut windows_core::PWSTR,
        ppszusertype: *mut windows_core::PWSTR,
        ppszverindprogid: *mut windows_core::PWSTR,
    ) -> windows_core::Result<()> {
        let (prog_id, user_type, ver_ind_prog_id) = self.0.get_class_details(unsafe { *clsid })?;

        opc_classic_utils::write_caller_allocated_ptr!(
            ppszprogid,
            opc_classic_utils::alloc_callee_wstring!(prog_id)?
        )?;
        opc_classic_utils::write_caller_allocated_ptr!(
            ppszusertype,
            opc_classic_utils::alloc_callee_wstring!(user_type)?
        )?;
        opc_classic_utils::write_caller_allocated_ptr!(
            ppszverindprogid,
            opc_classic_utils::alloc_callee_wstring!(ver_ind_prog_id)?
        )?;

        Ok(())
    }

    fn CLSIDFromProgID(
        &self,
        szprogid: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::GUID> {
        let prog_id = unsafe {
            opc_classic_utils::CallerAllocatedWString::from_pcwstr(*szprogid)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.clsid_from_prog_id(prog_id)
    }
}

#[windows::core::implement(opc_comn_bindings::IOPCShutdown)]
pub struct OPCShutdown<T>(T)
where
    T: 'static + traits::OPCShutdown;

impl<T: traits::OPCShutdown> opc_comn_bindings::IOPCShutdown_Impl for OPCShutdown_Impl<T> {
    fn ShutdownRequest(&self, szreason: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let reason = unsafe {
            opc_classic_utils::CallerAllocatedWString::from_pcwstr(*szreason)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.shutdown_request(reason)
    }
}

pub mod traits {
    pub trait OPCCommon {
        fn set_locale_id(&self, dwlcid: u32) -> windows_core::Result<()>;
        fn get_locale_id(&self) -> windows_core::Result<u32>;
        fn query_available_locale_ids(&self) -> windows_core::Result<Vec<u32>>;
        fn get_error_string(&self, dwerror: windows_core::HRESULT) -> windows_core::Result<String>;
        fn set_client_name(&self, szname: String) -> windows_core::Result<()>;
    }

    pub trait OPCEnumGUID {
        fn next(&self, celt: u32) -> windows_core::Result<Vec<windows_core::GUID>>;
        fn skip(&self, celt: u32) -> windows_core::Result<()>;
        fn reset(&self) -> windows_core::Result<()>;
        fn clone_enum(&self) -> windows_core::Result<opc_comn_bindings::IOPCEnumGUID>;
    }

    pub trait OPCServerList {
        fn enum_classes_of_categories(
            &self,
            implemented: Vec<windows_core::GUID>,
            required: Vec<windows_core::GUID>,
        ) -> windows_core::Result<windows::Win32::System::Com::IEnumGUID>;

        fn get_class_details(
            &self,
            clsid: windows_core::GUID,
        ) -> windows_core::Result<(String, String)>;

        fn clsid_from_prog_id(&self, prog_id: String) -> windows_core::Result<windows_core::GUID>;
    }

    pub trait OPCServerList2 {
        fn enum_classes_of_categories(
            &self,
            implemented: Vec<windows_core::GUID>,
            required: Vec<windows_core::GUID>,
        ) -> windows_core::Result<opc_comn_bindings::IOPCEnumGUID>;

        fn get_class_details(
            &self,
            clsid: windows_core::GUID,
        ) -> windows_core::Result<(String, String, String)>;

        fn clsid_from_prog_id(&self, prog_id: String) -> windows_core::Result<windows_core::GUID>;
    }

    pub trait OPCShutdown {
        fn shutdown_request(&self, reason: String) -> windows_core::Result<()>;
    }
}
