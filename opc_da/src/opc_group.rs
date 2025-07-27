use opc_classic_utils::*;
use opc_da_bindings::*;
use windows_core::*;

#[windows::core::implement(IOPCGroupStateMgt, IOPCGroupStateMgt2)]
pub struct OPCGroup<T>(T)
where
    T: 'static + traits::OPCGroup + traits::OPCGroupStateMgt2;

impl<T: traits::OPCGroup + traits::OPCGroupStateMgt2> IOPCGroupStateMgt_Impl for OPCGroup_Impl<T> {
    fn GetState(
        &self,
        pupdateRate: *mut u32,
        pactive: *mut BOOL,
        ppname: *mut PWSTR,
        ptimebias: *mut i32,
        ppercentdeadband: *mut f32,
        plcid: *mut u32,
        phclientgroup: *mut u32,
        phservergroup: *mut u32,
    ) -> Result<()> {
        let state = self.0.get_state()?;

        write_caller_allocated_ptr!(pupdateRate, state.update_rate)?;
        write_caller_allocated_ptr!(pactive, BOOL::from(state.active))?;
        write_caller_allocated_ptr!(ptimebias, state.time_bias)?;
        write_caller_allocated_ptr!(ppercentdeadband, state.percent_deadband)?;
        write_caller_allocated_ptr!(plcid, state.locale_id)?;
        write_caller_allocated_ptr!(phclientgroup, state.client_group)?;
        write_caller_allocated_ptr!(phservergroup, state.server_group)?;

        // Handle name separately as it's a string
        if !ppname.is_null() {
            let name_ptr = alloc_callee_wstring!(state.name)?;
            unsafe { *ppname = name_ptr };
        }

        Ok(())
    }

    fn SetState(
        &self,
        prequestedupdaterate: *const u32,
        previsedupdaterate: *mut u32,
        pactive: *const BOOL,
        ptimebias: *const i32,
        ppercentdeadband: *const f32,
        plcid: *const u32,
        phclientgroup: *const u32,
    ) -> Result<()> {
        let requested_update_rate = if !prequestedupdaterate.is_null() {
            Some(unsafe { *prequestedupdaterate })
        } else {
            None
        };

        let active = if !pactive.is_null() {
            Some(unsafe { (*pactive).as_bool() })
        } else {
            None
        };

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

        let locale_id = if !plcid.is_null() {
            Some(unsafe { *plcid })
        } else {
            None
        };

        let client_group = if !phclientgroup.is_null() {
            Some(unsafe { *phclientgroup })
        } else {
            None
        };

        let (
            revised_update_rate,
            new_active,
            new_time_bias,
            new_percent_deadband,
            new_locale_id,
            new_client_group,
        ) = self.0.set_state(
            requested_update_rate,
            active,
            time_bias,
            percent_deadband,
            locale_id,
            client_group,
        )?;

        write_caller_allocated_ptr!(previsedupdaterate, revised_update_rate)?;

        Ok(())
    }

    fn SetName(&self, szname: &PCWSTR) -> Result<()> {
        let name = unsafe {
            CallerAllocatedWString::from_pcwstr(*szname)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.set_name(name)
    }

    fn CloneGroup(&self, szname: &PCWSTR, riid: *const GUID) -> Result<IUnknown> {
        let name = unsafe {
            CallerAllocatedWString::from_pcwstr(*szname)
                .to_string()
                .ok_or(windows::Win32::Foundation::E_POINTER)
        }?;

        self.0.clone_group(name, riid)
    }
}

impl<T: traits::OPCGroup + traits::OPCGroupStateMgt2> IOPCGroupStateMgt2_Impl for OPCGroup_Impl<T> {
    fn SetKeepAlive(&self, dwkeepalivetime: u32) -> Result<u32> {
        self.0.set_keep_alive(dwkeepalivetime)
    }

    fn GetKeepAlive(&self) -> Result<u32> {
        self.0.get_keep_alive()
    }
}

pub mod traits {
    use super::*;

    #[derive(Debug, Clone)]
    pub struct GroupState {
        pub update_rate: u32,
        pub active: bool,
        pub name: String,
        pub time_bias: i32,
        pub percent_deadband: f32,
        pub locale_id: u32,
        pub client_group: u32,
        pub server_group: u32,
    }

    #[derive(Debug, Clone)]
    pub struct GroupState2 {
        pub update_rate: u32,
        pub active: bool,
        pub name: String,
        pub time_bias: i32,
        pub percent_deadband: f32,
        pub locale_id: u32,
        pub client_group: u32,
        pub server_group: u32,
        pub keep_alive: u32,
    }

    pub trait OPCGroup {
        fn get_state(&self) -> Result<GroupState>;

        fn set_state(
            &self,
            requested_update_rate: Option<u32>,
            active: Option<bool>,
            time_bias: Option<i32>,
            percent_deadband: Option<f32>,
            locale_id: Option<u32>,
            client_group: Option<u32>,
        ) -> Result<(u32, bool, i32, f32, u32, u32)>;

        fn set_name(&self, name: String) -> Result<()>;

        fn clone_group(&self, name: String, riid: *const GUID) -> Result<IUnknown>;
    }

    pub trait OPCGroupStateMgt2: OPCGroup {
        fn get_state2(&self) -> Result<GroupState2>;

        fn set_state2(
            &self,
            requested_update_rate: Option<u32>,
            active: Option<bool>,
            time_bias: Option<i32>,
            percent_deadband: Option<f32>,
            locale_id: Option<u32>,
            client_group: Option<u32>,
            keep_alive: Option<u32>,
        ) -> Result<(u32, bool, i32, f32, u32, u32, u32)>;

        fn set_keep_alive(&self, keep_alive_time: u32) -> Result<u32>;

        fn get_keep_alive(&self) -> Result<u32>;
    }
}
