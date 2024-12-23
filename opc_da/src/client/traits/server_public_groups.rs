use crate::utils::LocalPointer;

/// Server public groups management functionality.
///
/// Provides methods to access and manage public groups that can be shared
/// between multiple OPC clients for more efficient server resource usage.
pub trait ServerPublicGroupsTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCServerPublicGroups>;

    /// Gets a public group by its name.
    ///
    /// # Arguments
    /// * `name` - Name of the public group to retrieve
    /// * `id` - Interface ID for the requested group interface
    ///
    /// # Returns
    /// The requested interface pointer for the public group
    fn get_public_group_by_name(
        &self,
        name: &str,
        id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        let name = LocalPointer::from(name);

        unsafe { self.interface()?.GetPublicGroupByName(name.as_pcwstr(), id) }
    }

    /// Removes a public group from the server.
    ///
    /// # Arguments
    /// * `server_group` - Server handle of the group to remove
    /// * `force` - If true, removes group even if clients are connected
    ///
    /// # Returns
    /// Ok(()) if the group was successfully removed
    fn remove_public_group(&self, server_group: u32, force: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?
                .RemovePublicGroup(server_group, windows::Win32::Foundation::BOOL::from(force))
        }
    }
}
