use crate::client::memory::LocalPointer;
use opc_da_bindings::IOPCServerPublicGroups;
use std::str::FromStr;
use windows::Win32::Foundation::BOOL;

pub trait ServerPublicGroupsTrait {
    fn interface(&self) -> windows_core::Result<&IOPCServerPublicGroups>;

    fn get_public_group_by_name(
        &self,
        name: &str,
        id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        let name = LocalPointer::from_str(name)?;

        unsafe { self.interface()?.GetPublicGroupByName(name.as_pcwstr(), id) }
    }

    fn remove_public_group(&self, server_group: u32, force: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?
                .RemovePublicGroup(server_group, BOOL::from(force))
        }
    }
}
