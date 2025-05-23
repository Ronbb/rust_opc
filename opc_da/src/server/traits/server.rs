use crate::def::{EnumScope, ServerStatus};

use super::def::*;

/// Server trait
///
/// This trait defines the methods that a opc da server must implement to be used by the opc da client.
pub trait ServerTrait {
    /// Sets the locale ID to be used for string conversions.
    ///
    /// Implements the `IOPCCommon::SetLocaleID` method.
    ///
    /// # Arguments
    /// * `locale_id` - The locale ID to set
    ///
    /// # Returns
    /// An error if the locale ID is not supported
    fn set_locale_id(&self, locale_id: u32) -> windows::core::Result<()>;

    /// Returns the current locale ID.
    ///
    /// Implements the `IOPCCommon::GetLocaleID` method.
    ///
    /// # Returns
    /// The current locale ID
    fn get_locale_id(&self) -> windows::core::Result<u32>;

    /// Returns the list of available locale IDs.
    ///
    /// Implements the `IOPCCommon::QueryAvailableLocaleIDs` method.
    ///
    /// # Returns
    /// The list of available locale IDs
    fn query_available_locale_ids(&self) -> windows::core::Result<Vec<u32>>;

    /// Returns the error string for the given error code.
    ///
    /// Implements the `IOPCCommon::GetErrorString` method.
    ///
    /// # Arguments
    /// * `error` - The error code
    ///
    /// # Returns
    /// The error string for the given error code
    fn get_error_string(&self, error: i32) -> windows::core::Result<String>;

    /// Sets the name of the client application.  
    ///  
    /// Implements the `IOPCCommon::SetClientName` method.  
    ///  
    /// # Arguments  
    /// * `name` - The client application name
    ///  
    /// # Returns  
    /// An error if the operation fails
    fn set_client_name(&self, name: String) -> windows::core::Result<()>;

    /// Returns the name of the client application.
    ///
    /// Implements the `IConnectionPointContainer::EnumConnectionPoints` method.
    ///
    /// # Returns
    ///
    /// A result containing a vector of connection points
    fn enum_connection_points(
        &self,
    ) -> windows::core::Result<Vec<windows::Win32::System::Com::IConnectionPoint>>;

    /// Returns the connection point for the given reference interface ID.
    ///
    /// Implements the `IConnectionPointContainer::FindConnectionPoint` method.
    ///
    /// # Arguments
    /// * `reference_interface_id` - The reference interface ID
    ///
    /// # Returns
    /// The connection point
    fn find_connection_point(
        &self,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint>;

    /// Returns the list of available properties for the given item ID.
    ///
    /// Implements the `IOPCBrowse::QueryAvailableProperties` method.
    ///
    /// # Arguments
    /// * `item_id` - The item ID
    ///
    /// # Returns
    /// The list of available properties
    fn query_available_properties(
        &self,
        item_id: String,
    ) -> windows::core::Result<Vec<AvailableProperty>>;

    /// Returns the properties for the given item ID.
    ///
    /// Implements the `IOPCItemProperties::GetItemProperties` method.
    ///
    /// # Arguments
    /// * `item_id` - The item ID
    /// * `property_ids` - The list of property IDs
    ///
    /// # Returns
    /// The properties for the given item ID
    fn get_item_properties(
        &self,
        item_id: String,
        property_ids: Vec<u32>,
    ) -> windows::core::Result<Vec<ItemPropertyData>>;

    /// Lookup the item IDs for the given item ID and property IDs.
    ///
    /// Implements the `IOPCBrowse::LookupItemIDs` method.
    ///
    /// # Arguments
    /// * `item_id` - The item ID
    /// * `property_ids` - The list of property IDs
    ///
    /// # Returns
    /// The item IDs for the given item ID and property IDs
    fn lookup_item_ids(
        &self,
        item_id: String,
        property_ids: Vec<u32>,
    ) -> windows::core::Result<Vec<NewItem>>;

    /// Returns the properties for the given item IDs.
    ///
    /// Implements the `IOPCItemProperties::GetProperties` method.
    ///
    /// # Arguments
    /// * `item_ids` - The list of item IDs
    /// * `return_property_values` - Whether to return property values
    /// * `property_ids` - The list of property IDs
    ///
    /// # Returns
    /// The properties for the given item IDs
    fn get_properties(
        &self,
        item_ids: Vec<String>,
        return_property_values: bool,
        property_ids: Vec<u32>,
    ) -> windows::core::Result<Vec<ItemProperties>>;

    /// Browse the server for items.
    ///
    /// Implements the `IOPCBrowse::Browse` method.
    ///
    /// # Arguments
    /// * `item_id` - The item ID
    /// * `continuation_point` - The continuation point
    /// * `max_elements_returned` - The maximum number of elements to return
    /// * `browse_filter` - The browse filter
    /// * `element_name_filter` - The element name filter
    /// * `vendor_filter` - The vendor filter
    /// * `return_all_properties` - Whether to return all properties
    /// * `return_property_values` - Whether to return property values
    /// * `property_ids` - The list of property IDs
    ///
    /// # Returns
    /// The browse result
    #[allow(clippy::too_many_arguments)]
    fn browse(
        &self,
        item_id: String,
        continuation_point: Option<String>,
        max_elements_returned: u32,
        browse_filter: BrowseFilter,
        element_name_filter: String,
        vendor_filter: String,
        return_all_properties: bool,
        return_property_values: bool,
        property_ids: Vec<u32>,
    ) -> windows::core::Result<BrowseResult>;

    /// Get the public group by name.
    ///
    /// Implements the `IOPCServerPublicGroups::GetPublicGroupByName` method.
    ///
    /// # Arguments
    /// * `name` - The name of the public group
    /// * `reference_interface_id` - The reference interface ID
    ///
    /// # Returns
    /// The public group
    fn get_public_group_by_name(
        &self,
        name: String,
        reference_interface_id: u128,
    ) -> windows::core::Result<windows::core::IUnknown>;

    fn remove_public_group(&self, server_group: u32, force: bool) -> windows::core::Result<()>;

    fn query_organization(&self) -> windows::core::Result<NamespaceType>;

    fn change_browse_position(
        &self,
        browse_direction: BrowseDirection,
    ) -> windows::core::Result<()>;

    fn browse_opc_item_ids(
        &self,
        browse_filter_type: BrowseType,
        filter_criteria: String,
        variant_data_type_filter: u16,
        access_rights_filter: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString>;

    fn get_item_id(&self, item_data_id: String) -> windows::core::Result<String>;

    fn browse_access_paths(&self, item_id: String) -> windows::core::Result<Vec<String>>;

    fn read(&self, items: Vec<ItemWithMaxAge>) -> windows::core::Result<Vec<VqtWithError>>;

    fn write_vqt(
        &self,
        items: Vec<ItemOptionalVqt>,
    ) -> windows::core::Result<Vec<windows::core::HRESULT>>;

    #[allow(clippy::too_many_arguments)]
    fn add_group(
        &self,
        name: String,
        active: bool,
        requested_update_rate: u32,
        client_group: u32,
        time_bias: Option<i32>,
        percent_deadband: Option<f32>,
        locale_id: u32,
        reference_interface_id: Option<u128>,
    ) -> windows::core::Result<GroupInfo>;

    fn get_error_string_locale(&self, error: i32, locale: u32) -> windows::core::Result<String>;

    fn get_group_by_name(
        &self,
        name: String,
        reference_interface_id: Option<u128>,
    ) -> windows::core::Result<windows::core::IUnknown>;

    fn get_status(&self) -> windows::core::Result<ServerStatus>;

    fn remove_group(&self, server_group: u32, force: bool) -> windows::core::Result<()>;

    fn create_group_enumerator(
        &self,
        scope: EnumScope,
        reference_interface_id: Option<u128>,
    ) -> windows::core::Result<windows::core::IUnknown>;
}
