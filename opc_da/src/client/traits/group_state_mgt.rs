use std::str::FromStr;

use crate::{
    client::{memory::LocalPointer, RemotePointer},
    def,
};

/// Group state management functionality.
///
/// Provides methods to get and set various group state parameters including:
/// - Update rate
/// - Active state
/// - Time bias
/// - Deadband
/// - Locale ID
/// - Group handles
pub trait GroupStateMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt>;

    /// Gets the current state of the group.
    ///
    /// Returns a GroupState structure containing all group parameters.
    fn get_state(&self) -> windows::core::Result<def::GroupState> {
        let mut state = def::GroupState::default();
        let mut active = windows::Win32::Foundation::BOOL::default();
        let name = {
            let mut name = RemotePointer::new();
            unsafe {
                self.interface()?.GetState(
                    &mut state.update_rate,
                    &mut active,
                    name.as_mut_pwstr_ptr(),
                    &mut state.time_bias,
                    &mut state.percent_deadband,
                    &mut state.locale_id,
                    &mut state.client_handle,
                    &mut state.server_handle,
                )?;
            }
            name
        };

        state.active = active.as_bool();
        state.name = name.try_into()?;

        Ok(state)
    }

    /// Sets one or more group state parameters.
    ///
    /// # Arguments
    /// * `update_rate` - Requested group update rate in milliseconds
    /// * `active` - Group active state
    /// * `time_bias` - Time bias from UTC in minutes
    /// * `percent_deadband` - Percent deadband for analog items
    /// * `locale_id` - Locale ID for status/error strings
    /// * `client_group_handle` - Client-provided handle
    ///
    /// # Returns
    /// The actual update rate set by the server
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

    /// Sets the name of the group.
    fn set_name(&self, name: &str) -> windows::core::Result<()> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.SetName(name.as_pwstr()) }
    }

    /// Creates a copy of the group with a new name.
    ///
    /// # Arguments
    /// * `name` - Name for the new group
    /// * `id` - Client-provided GUID for the new group
    fn clone_group(
        &self,
        name: &str,
        id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.CloneGroup(name.as_pwstr(), id) }
    }
}
