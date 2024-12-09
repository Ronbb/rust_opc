use crate::client::memory::LocalPointer;
use std::str::FromStr;

pub trait ServerPublicGroupsTrait {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCServerPublicGroups>;

    fn get_public_group_by_name(
        &self,
        name: &str,
        id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.GetPublicGroupByName(name.as_pcwstr(), id) }
    }

    fn remove_public_group(&self, server_group: u32, force: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?
                .RemovePublicGroup(server_group, windows::Win32::Foundation::BOOL::from(force))
        }
    }
}
