use windows::core::Interface as _;

use crate::{
    client::{GroupIterator, StringIterator},
    utils::{LocalPointer, RemotePointer},
};

/// OPC Server management functionality.
///
/// Provides methods to create and manage groups within an OPC server,
/// as well as monitor server status and enumerate existing groups.
pub trait ServerTrait<Group: TryFrom<windows::core::IUnknown, Error = windows::core::Error>> {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCServer>;

    /// Adds a new group to the OPC server.
    ///
    /// # Arguments
    /// * `name` - Group name for identification
    /// * `active` - Whether the group should initially be active
    /// * `client_handle` - Client-assigned handle for the group
    /// * `update_rate` - Requested update rate in milliseconds
    /// * `locale_id` - Locale ID for text strings
    /// * `time_bias` - Time zone bias in minutes from UTC
    /// * `percent_deadband` - Percent change required to trigger updates
    ///
    /// # Returns
    /// The newly created group object
    ///
    /// # Errors
    /// Returns E_POINTER if group creation fails
    #[allow(clippy::too_many_arguments)]
    fn add_group(
        &self,
        name: &str,
        active: bool,
        client_handle: u32,
        update_rate: u32,
        locale_id: u32,
        time_bias: i32,
        percent_deadband: f32,
        revised_percent_deadband: &mut u32,
        server_handle: &mut u32,
    ) -> windows::core::Result<Group> {
        let mut group = None;
        let group_name = LocalPointer::from(name);
        let group_name = group_name.as_pcwstr();

        unsafe {
            self.interface()?.AddGroup(
                group_name,
                active,
                update_rate,
                client_handle,
                &time_bias,
                &percent_deadband,
                locale_id,
                server_handle,
                revised_percent_deadband,
                &opc_da_bindings::IOPCItemMgt::IID,
                &mut group,
            )?;
        }

        match group {
            None => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Failed to add group, returned null",
            )),
            Some(group) => group.cast::<windows::core::IUnknown>()?.try_into(),
        }
    }

    /// Gets the current server status.
    ///
    /// # Returns
    /// Server status structure containing vendor info, time, state,
    /// and group counts
    fn get_status(
        &self,
    ) -> windows::core::Result<RemotePointer<opc_da_bindings::tagOPCSERVERSTATUS>> {
        let status = unsafe { self.interface()?.GetStatus()? };
        Ok(RemotePointer::from_raw(status))
    }

    /// Removes a group from the server.
    ///
    /// # Arguments
    /// * `server_handle` - Server's handle for the group
    /// * `force` - If true, remove even if clients are connected
    fn remove_group(&self, server_handle: u32, force: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?.RemoveGroup(server_handle, force)?;
        }
        Ok(())
    }

    /// Creates an enumerator for groups.
    ///
    /// # Arguments
    /// * `scope` - Scope of groups to enumerate (public, private, or all)
    ///
    /// # Returns
    /// Enumerator interface for iterating through groups
    fn create_group_enumerator(
        &self,
        scope: opc_da_bindings::tagOPCENUMSCOPE,
    ) -> windows::core::Result<GroupIterator<Group>> {
        let enumerator = unsafe {
            self.interface()?
                .CreateGroupEnumerator(scope, &windows::Win32::System::Com::IEnumUnknown::IID)?
        };

        Ok(GroupIterator::new(enumerator.cast()?))
    }

    /// Creates an enumerator for group names.
    ///
    /// # Arguments
    /// * `scope` - Scope of group names to enumerate (public, private, or all)
    ///
    /// # Returns
    /// Enumerator interface for iterating through group names
    fn create_group_name_enumerator(
        &self,
        scope: opc_da_bindings::tagOPCENUMSCOPE,
    ) -> windows::core::Result<StringIterator> {
        let enumerator = unsafe {
            self.interface()?
                .CreateGroupEnumerator(scope, &windows::Win32::System::Com::IEnumString::IID)?
        };

        Ok(StringIterator::new(enumerator.cast()?))
    }
}
