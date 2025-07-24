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

pub mod traits {
    pub trait OPCCommon {
        fn set_locale_id(&self, dwlcid: u32) -> windows_core::Result<()>;
        fn get_locale_id(&self) -> windows_core::Result<u32>;
        fn query_available_locale_ids(&self) -> windows_core::Result<Vec<u32>>;
        fn get_error_string(&self, dwerror: windows_core::HRESULT) -> windows_core::Result<String>;
        fn set_client_name(&self, szname: String) -> windows_core::Result<()>;
    }
}
