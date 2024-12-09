use std::str::FromStr;

use crate::client::memory::{LocalPointer, RemotePointer};
use opc_da_bindings::{
    tagOPCBROWSEDIRECTION, tagOPCBROWSETYPE, tagOPCNAMESPACETYPE, IOPCBrowseServerAddressSpace,
};

pub trait BrowseServerAddressSpaceTrait {
    fn interface(&self) -> windows::core::Result<&IOPCBrowseServerAddressSpace>;

    fn query_organization(&self) -> windows::core::Result<tagOPCNAMESPACETYPE> {
        unsafe { self.interface()?.QueryOrganization() }
    }

    fn change_browse_position(
        &self,
        browse_direction: tagOPCBROWSEDIRECTION,
        position: &str,
    ) -> windows::core::Result<()> {
        let position = LocalPointer::from_str(position)?;

        unsafe {
            self.interface()?
                .ChangeBrowsePosition(browse_direction, position.as_pwstr())
        }
    }

    fn browse_opc_item_ids(
        &self,
        browse_type: tagOPCBROWSETYPE,
        filter_criteria: &str,
        datatype_filter: u16,
        access_rights_filter: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString> {
        let filter_criteria = LocalPointer::from_str(filter_criteria)?;

        unsafe {
            self.interface()?.BrowseOPCItemIDs(
                browse_type,
                filter_criteria.as_pwstr(),
                datatype_filter,
                access_rights_filter,
            )
        }
    }

    fn get_item_id(&self, item_data_id: &str) -> windows::core::Result<String> {
        let item_data_id = LocalPointer::from_str(item_data_id)?;

        let output = unsafe { self.interface()?.GetItemID(item_data_id.as_pwstr())? };

        RemotePointer::from(output).try_into()
    }

    fn browse_access_paths(
        &self,
        item_id: &str,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString> {
        let item_id = LocalPointer::from_str(item_id)?;
        unsafe { self.interface()?.BrowseAccessPaths(item_id.as_pwstr()) }
    }
}
