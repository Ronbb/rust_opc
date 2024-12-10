use windows::core::Interface as _;

use crate::client::LocalPointer;

pub trait ServerTrait<Group>
where
    Group: TryFrom<windows::core::IUnknown, Error = windows::core::Error>,
{
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCServer>;

    /// Adds a new group to the OPC server.  
    ///
    /// # Safety  
    /// This method is unsafe because it calls into COM interfaces.  
    /// The caller must ensure that the COM server is properly initialized.  
    ///
    /// # Arguments  
    /// * `name` - The name of the group  
    /// * `active` - Whether the group should be active  
    /// * `client_handle` - The client handle for the group
    /// * `update_rate` - The update rate for the group
    /// * `locale_id` - The locale id for the group
    /// * `time_bias` - The time bias for the group
    /// * `percent_deadband` - The percent deadband for the group
    ///
    /// # Returns
    /// The newly created group
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
    ) -> windows::core::Result<Group> {
        let mut group = None;
        let mut group_server_handle = 0u32;
        let group_name = LocalPointer::from(name);
        let group_name = group_name.as_pcwstr();
        let mut revised_percent_deadband = 0;

        unsafe {
            self.interface()?.AddGroup(
                group_name,
                windows::Win32::Foundation::BOOL::from(active),
                update_rate,
                client_handle,
                &time_bias,
                &percent_deadband,
                locale_id,
                &mut group_server_handle,
                &mut revised_percent_deadband,
                &opc_da_bindings::IOPCItemMgt::IID,
                &mut group,
            )?;

            match group {
                None => Err(windows::core::Error::new(
                    windows::Win32::Foundation::E_POINTER,
                    "Failed to add group, returned null",
                )),
                Some(group) => group.try_into(),
            }
        }
    }

    fn get_status(&self) -> windows::core::Result<opc_da_bindings::tagOPCSERVERSTATUS> {
        let status = unsafe { self.interface()?.GetStatus()?.as_ref() };
        match status {
            Some(status) => Ok(*status),
            None => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_FAIL,
                "Failed to get server status",
            )),
        }
    }

    fn remove_group(&self, server_handle: u32, force: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?
                .RemoveGroup(server_handle, windows::Win32::Foundation::BOOL::from(force))?;
        }
        Ok(())
    }

    fn create_group_enumerator(
        &self,
        scope: opc_da_bindings::tagOPCENUMSCOPE,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumUnknown> {
        let enumerator = unsafe {
            self.interface()?
                .CreateGroupEnumerator(scope, &windows::Win32::System::Com::IEnumUnknown::IID)?
        };

        enumerator.cast()
    }
}
