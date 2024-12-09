pub trait GroupStateMgt2Trait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt2>;

    fn set_keep_alive(&self, keep_alive_time: u32) -> windows::core::Result<u32> {
        unsafe { self.interface()?.SetKeepAlive(keep_alive_time) }
    }

    fn get_keep_alive(&self) -> windows::core::Result<u32> {
        unsafe { self.interface()?.GetKeepAlive() }
    }
}
