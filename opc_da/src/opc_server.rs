use opc_classic_utils::*;
use opc_da_bindings::*;
use windows_core::*;

#[windows::core::implement(IOPCServer)]
pub struct OPCServer<T>(T)
where
    T: 'static + traits::OPCServer;

impl<T: traits::OPCServer> IOPCServer_Impl for OPCServer_Impl<T> {
    fn AddGroup(
        &self,
        szname: &PCWSTR,
        bactive: BOOL,
        dwrequestedupdaterate: u32,
        hclientgroup: u32,
        ptimebias: *const i32,
        ppercentdeadband: *const f32,
        dwlcid: u32,
        phservergroup: *mut u32,
        previsedupdaterate: *mut u32,
        riid: *const GUID,
        ppunk: OutRef<'_, IUnknown>,
    ) -> Result<()> {
        let name = unsafe {
            CallerAllocatedWString::from_pcwstr(*szname)
                .to_string_lossy()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        let time_bias = if !ptimebias.is_null() {
            Some(unsafe { *ptimebias })
        } else {
            None
        };

        let percent_deadband = if !ppercentdeadband.is_null() {
            Some(unsafe { *ppercentdeadband })
        } else {
            None
        };

        let (server_group, revised_update_rate, group_interface) = self.0.add_group(
            name,
            bactive.as_bool(),
            dwrequestedupdaterate,
            hclientgroup,
            time_bias,
            percent_deadband,
            dwlcid,
        )?;

        write_caller_allocated_ptr!(phservergroup, server_group)?;
        write_caller_allocated_ptr!(previsedupdaterate, revised_update_rate)?;
        ppunk.write(Some(group_interface))?;

        Ok(())
    }

    fn GetErrorString(&self, dwerror: HRESULT, dwlocale: u32) -> Result<PWSTR> {
        alloc_callee_wstring!(self.0.get_error_string(dwerror, dwlocale)?)
    }

    fn GetGroupByName(&self, szname: &PCWSTR, riid: *const GUID) -> Result<IUnknown> {
        let name = unsafe {
            CallerAllocatedWString::from_pcwstr(*szname)
                .to_string_lossy()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.get_group_by_name(name, riid)
    }

    fn GetStatus(&self) -> Result<*mut tagOPCSERVERSTATUS> {
        let status = self.0.get_status()?;
        let ptr = CalleeAllocatedPtr::from_value(&status)?;
        Ok(ptr.into_raw())
    }

    fn RemoveGroup(&self, hservergroup: u32, bforce: BOOL) -> Result<()> {
        self.0.remove_group(hservergroup, bforce.as_bool())
    }

    fn CreateGroupEnumerator(
        &self,
        dwscope: tagOPCENUMSCOPE,
        riid: *const GUID,
    ) -> Result<IUnknown> {
        self.0.create_group_enumerator(dwscope, riid)
    }
}

pub mod traits {
    use super::*;
    use opc_da_bindings::tagOPCSERVERSTATUS;

    pub trait OPCServer {
        fn add_group(
            &self,
            name: String,
            active: bool,
            requested_update_rate: u32,
            client_group: u32,
            time_bias: Option<i32>,
            percent_deadband: Option<f32>,
            locale_id: u32,
        ) -> Result<(u32, u32, IUnknown)>;

        fn get_error_string(&self, error: HRESULT, locale: u32) -> Result<String>;

        fn get_group_by_name(&self, name: String, riid: *const GUID) -> Result<IUnknown>;

        fn get_status(&self) -> Result<tagOPCSERVERSTATUS>;

        fn remove_group(&self, server_group: u32, force: bool) -> Result<()>;

        fn create_group_enumerator(
            &self,
            scope: tagOPCENUMSCOPE,
            riid: *const GUID,
        ) -> Result<IUnknown>;
    }
}
