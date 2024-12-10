pub trait PublicGroupStateMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCPublicGroupStateMgt>;

    fn get_state(&self) -> windows::core::Result<bool> {
        unsafe { self.interface()?.GetState() }.map(|v| v.as_bool())
    }

    fn move_to_public(&self) -> windows::core::Result<()> {
        unsafe { self.interface()?.MoveToPublic() }
    }
}
