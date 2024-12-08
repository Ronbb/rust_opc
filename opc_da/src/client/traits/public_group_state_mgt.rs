pub trait PublicGroupStateMgtTrait {
    fn public_group_state_mgt(&self) -> &opc_da_bindings::IOPCPublicGroupStateMgt;

    fn get_state(&self) -> windows::core::Result<bool> {
        unsafe { self.public_group_state_mgt().GetState() }.map(|v| v.as_bool())
    }

    fn move_to_public(&self) -> windows::core::Result<()> {
        unsafe { self.public_group_state_mgt().MoveToPublic() }
    }
}
