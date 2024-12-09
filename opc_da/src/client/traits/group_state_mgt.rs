use std::str::FromStr;

use crate::{client::memory::LocalPointer, def};

pub trait GroupStateMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt>;

    fn get_state(&self) -> windows::core::Result<def::GroupState> {
        let mut state = def::GroupState::default();

        let mut active = windows::Win32::Foundation::BOOL::default();
        let mut name = windows::core::PWSTR::null();

        unsafe {
            self.interface()?.GetState(
                &mut state.update_rate,
                &mut active,
                &mut name,
                &mut state.time_bias,
                &mut state.percent_deadband,
                &mut state.locale_id,
                &mut state.client_group_handle,
                &mut state.server_group_handle,
            )
        }?;

        state.active = active.as_bool();
        state.name = unsafe { name.to_string() }?;

        Ok(state)
    }

    fn set_state(
        &self,
        update_rate: Option<u32>,
        active: Option<bool>,
        time_bias: Option<i32>,
        percent_deadband: Option<f32>,
        locale_id: Option<u32>,
        client_group_handle: Option<u32>,
    ) -> windows::core::Result<u32> {
        let requested_update_rate = LocalPointer::new(update_rate);
        let mut revised_update_rate = LocalPointer::new(Some(0));
        let active = LocalPointer::new(active.map(windows::Win32::Foundation::BOOL::from));
        let time_bias = LocalPointer::new(time_bias);
        let percent_deadband = LocalPointer::new(percent_deadband);
        let locale_id = LocalPointer::new(locale_id);
        let client_group_handle = LocalPointer::new(client_group_handle);

        unsafe {
            self.interface()?.SetState(
                requested_update_rate.as_ptr(),
                revised_update_rate.as_mut_ptr(),
                active.as_ptr(),
                time_bias.as_ptr(),
                percent_deadband.as_ptr(),
                locale_id.as_ptr(),
                client_group_handle.as_ptr(),
            )
        }?;

        Ok(revised_update_rate.into_inner().unwrap_or_default())
    }

    fn set_name(&self, name: &str) -> windows::core::Result<()> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.SetName(name.as_pwstr()) }
    }

    fn clone_group(
        &self,
        name: &str,
        id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.CloneGroup(name.as_pwstr(), id) }
    }
}
