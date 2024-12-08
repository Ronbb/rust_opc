use crate::client::{memory::LocalPointer, traits::group_state_mgt::GroupStateMgtTrait};

pub trait GroupStateMgt2Trait: GroupStateMgtTrait {
    fn group_state_mgt2(&self) -> &opc_da_bindings::IOPCGroupStateMgt2;

    fn set_keep_alive(&self, keep_alive_time: u32) -> windows_core::Result<u32> {
        unsafe { self.group_state_mgt2().SetKeepAlive(keep_alive_time) }
    }

    fn get_keep_alive(&self) -> windows_core::Result<u32> {
        unsafe { self.group_state_mgt2().GetKeepAlive() }
    }
}
