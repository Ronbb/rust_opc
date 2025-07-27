#[windows::core::implement(opc_comn_bindings::IOPCCommon, opc_comn_bindings::IOPCShutdown)]
pub struct OpcServerBase<T>(T)
where
    T: 'static + super::traits::opc_common::OpcCommon + super::traits::opc_shutdown::OpcShutdown;

impl<T: super::traits::opc_common::OpcCommon + super::traits::opc_shutdown::OpcShutdown>
    opc_comn_bindings::IOPCCommon_Impl for OpcServerBase_Impl<T>
{
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
                .to_string_lossy()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?)
    }
}

impl<T: super::traits::opc_common::OpcCommon + super::traits::opc_shutdown::OpcShutdown>
    opc_comn_bindings::IOPCShutdown_Impl for OpcServerBase_Impl<T>
{
    fn ShutdownRequest(&self, szreason: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let reason = unsafe {
            opc_classic_utils::CallerAllocatedWString::from_pcwstr(*szreason)
                .to_string_lossy()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.shutdown_request(reason)
    }
}
