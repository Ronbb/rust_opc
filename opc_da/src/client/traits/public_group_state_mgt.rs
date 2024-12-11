/// Public group state management functionality.
///
/// Provides methods to manage public groups in OPC servers. Public groups
/// can be shared between multiple clients, allowing for more efficient
/// server resource usage.
pub trait PublicGroupStateMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCPublicGroupStateMgt>;

    /// Gets the public state of the group.
    ///
    /// # Returns
    /// `true` if the group is public, `false` if it is private
    fn get_state(&self) -> windows::core::Result<bool> {
        unsafe { self.interface()?.GetState() }.map(|v| v.as_bool())
    }

    /// Converts a private group to a public group.
    ///
    /// # Returns
    /// Ok(()) if the group was successfully converted to public
    ///
    /// # Notes
    /// Once a group becomes public, it remains public until the server
    /// is shut down or the group is deleted.
    fn move_to_public(&self) -> windows::core::Result<()> {
        unsafe { self.interface()?.MoveToPublic() }
    }
}
