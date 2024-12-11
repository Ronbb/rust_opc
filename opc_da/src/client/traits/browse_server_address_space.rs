use std::str::FromStr;

use crate::client::memory::{LocalPointer, RemotePointer};
use opc_da_bindings::{
    tagOPCBROWSEDIRECTION, tagOPCBROWSETYPE, tagOPCNAMESPACETYPE, IOPCBrowseServerAddressSpace,
};

/// Server address space browsing functionality.
///
/// Provides methods to navigate and discover items in the OPC server's address space.
/// Supports hierarchical and flat address spaces with filtering capabilities.
pub trait BrowseServerAddressSpaceTrait {
    fn interface(&self) -> windows::core::Result<&IOPCBrowseServerAddressSpace>;

    /// Queries the organization type of the server's address space.
    ///
    /// # Returns
    /// The namespace type (hierarchical or flat)
    fn query_organization(&self) -> windows::core::Result<tagOPCNAMESPACETYPE> {
        unsafe { self.interface()?.QueryOrganization() }
    }

    /// Changes the current position in the server's address space.
    ///
    /// # Arguments
    /// * `browse_direction` - Direction to move (up, down, or to)
    /// * `position` - Target position string (branch name for down/to)
    ///
    /// # Returns
    /// Result indicating success or failure of position change
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

    /// Browses available item IDs at the current position.
    ///
    /// # Arguments
    /// * `browse_type` - Type of items to browse (leaf, branch, or flat)
    /// * `filter_criteria` - Pattern for filtering items
    /// * `datatype_filter` - VT_* type to filter by (0 for all)
    /// * `access_rights_filter` - Required access rights
    ///
    /// # Returns
    /// Enumerator for matching item IDs
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

    /// Gets fully qualified item ID from a leaf item.
    ///
    /// # Arguments
    /// * `item_data_id` - Item name at current position
    ///
    /// # Returns
    /// Fully qualified item ID string
    fn get_item_id(&self, item_data_id: &str) -> windows::core::Result<String> {
        let item_data_id = LocalPointer::from_str(item_data_id)?;

        let output = unsafe { self.interface()?.GetItemID(item_data_id.as_pwstr())? };

        RemotePointer::from(output).try_into()
    }

    /// Browses available access paths for an item.
    ///
    /// # Arguments
    /// * `item_id` - Fully qualified item ID
    ///
    /// # Returns
    /// Enumerator for available access paths
    fn browse_access_paths(
        &self,
        item_id: &str,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString> {
        let item_id = LocalPointer::from_str(item_id)?;
        unsafe { self.interface()?.BrowseAccessPaths(item_id.as_pwstr()) }
    }
}
