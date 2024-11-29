use crate::com::base::{SystemTime, Variant};

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
    fn set_locale_id(&self, locale_id: u32) -> windows_core::Result<()>;

    /// Returns the current locale ID.
    ///
    /// Implements the `IOPCCommon::GetLocaleID` method.
    ///
    /// # Returns
    /// The current locale ID
    fn get_locale_id(&self) -> windows_core::Result<u32>;

    /// Returns the list of available locale IDs.
    ///
    /// Implements the `IOPCCommon::QueryAvailableLocaleIDs` method.
    ///
    /// # Returns
    /// The list of available locale IDs
    fn query_available_locale_ids(&self) -> windows_core::Result<Vec<u32>>;

    /// Returns the error string for the given error code.
    ///
    /// Implements the `IOPCCommon::GetErrorString` method.
    ///
    /// # Arguments
    /// * `error` - The error code
    ///
    /// # Returns
    /// The error string for the given error code
    fn get_error_string(&self, error: i32) -> windows_core::Result<String>;

    /// Sets the name of the client application.  
    ///  
    /// Implements the `IOPCCommon::SetClientName` method.  
    ///  
    /// # Arguments  
    /// * `name` - The client application name (must not be empty and should be reasonable length)  
    ///  
    /// # Returns  
    /// An error if the name is empty or exceeds maximum length  
    fn set_client_name(&self, name: String) -> windows_core::Result<()>;

    /// Returns the name of the client application.
    ///
    /// Implements the `IConnectionPointContainer::EnumConnectionPoints` method.
    ///
    /// # Returns
    ///
    /// A list of connection points
    fn enum_connection_points(
        &self,
    ) -> windows_core::Result<Vec<windows::Win32::System::Com::IConnectionPoint>>;

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
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows::Win32::System::Com::IConnectionPoint>;

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
    ) -> windows_core::Result<Vec<AvailableProperty>>;

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
    ) -> windows_core::Result<Vec<ItemPropertyData>>;

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
    ) -> windows_core::Result<Vec<NewItem>>;

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
    ) -> windows_core::Result<Vec<ItemProperties>>;

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
    ) -> windows_core::Result<BrowseResult>;

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
    ) -> windows_core::Result<windows_core::IUnknown>;

    fn remove_public_group(&self, server_group: u32, force: bool) -> windows_core::Result<()>;

    fn query_organization(&self) -> windows_core::Result<NamespaceType>;

    fn change_browse_position(&self, browse_direction: BrowseDirection)
        -> windows_core::Result<()>;

    fn browse_opc_item_ids(
        &self,
        browse_filter_type: BrowseType,
        filter_criteria: String,
        variant_data_type_filter: u16,
        access_rights_filter: u32,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumString>;

    fn get_item_id(&self, item_data_id: String) -> windows_core::Result<String>;

    fn browse_access_paths(&self, item_id: String) -> windows_core::Result<Vec<String>>;

    fn read(&self, items: Vec<ItemWithMaxAge>) -> windows_core::Result<Vec<VqtWithError>>;

    fn write_vqt(
        &self,
        items: Vec<ItemOptionalVqt>,
    ) -> windows_core::Result<Vec<windows_core::HRESULT>>;

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
    ) -> windows_core::Result<GroupInfo>;

    fn get_error_string_locale(&self, error: i32, locale: u32) -> windows_core::Result<String>;

    fn get_group_by_name(
        &self,
        name: String,
        reference_interface_id: Option<u128>,
    ) -> windows_core::Result<windows_core::IUnknown>;

    fn get_status(&self) -> windows_core::Result<ServerStatus>;

    fn remove_group(&self, server_group: u32, force: bool) -> windows_core::Result<()>;

    fn create_group_enumerator(
        &self,
        scope: EnumScope,
        reference_interface_id: Option<u128>,
    ) -> windows_core::Result<windows_core::IUnknown>;
}

pub struct AvailableProperty {
    pub property_id: u32,
    pub description: String,
    pub data_type: u16,
}

pub struct ItemPropertyData {
    pub property_id: u32,
    pub data: Variant,
    pub error: windows_core::HRESULT,
}

pub struct NewItem {
    pub new_item_id: String,
    pub error: windows_core::HRESULT,
}

pub struct ItemProperties {
    pub error_id: windows_core::HRESULT,
    pub item_properties: Vec<ItemProperty>,
}

pub struct ItemProperty {
    pub data_type: u16,
    pub property_id: u32,
    pub item_id: String,
    pub description: String,
    pub value: Variant,
    pub error_id: windows_core::HRESULT,
}

pub enum BrowseFilter {
    All,
    Branches,
    Items,
}

pub struct BrowseResult {
    pub more_elements: bool,
    pub continuation_point: Option<String>,
    pub elements: Vec<BrowseElement>,
}

pub struct BrowseElement {
    pub name: String,
    pub item_id: String,
    pub flag_value: u32,
    pub item_properties: ItemProperties,
}

pub enum NamespaceType {
    Flat,
    Hierarchical,
}

pub enum BrowseDirection {
    Up,
    Down,
    To(String),
}

pub enum BrowseType {
    Branch,
    Leaf,
    Flat,
}

pub struct ItemWithMaxAge {
    pub item_id: String,
    pub max_age: u32,
}

pub struct Vqt {
    pub value: Variant,
    pub quality: u16,
    pub timestamp: SystemTime,
}

pub struct OptionalVqt {
    pub value: Variant,
    pub quality: Option<u16>,
    pub timestamp: Option<SystemTime>,
}

pub struct VqtWithError {
    pub value: Variant,
    pub quality: u16,
    pub timestamp: SystemTime,
    pub error: windows_core::HRESULT,
}

pub struct ItemOptionalVqt {
    pub item_id: String,
    pub optional_vqt: OptionalVqt,
}

pub struct GroupInfo {
    pub server_group: u32,
    pub revised_update_rate: u32,
    pub unknown: windows_core::IUnknown,
}

pub struct ServerStatus {
    pub start_time: SystemTime,
    pub current_time: SystemTime,
    pub last_update_time: SystemTime,
    pub server_state: ServerState,
    pub group_count: u32,
    pub band_width: u32,
    pub major_version: u16,
    pub minor_version: u16,
    pub build_number: u16,
    pub vendor_info: String,
}

pub enum ServerState {
    Running,
    Failed,
    NoConfig,
    Suspended,
    Test,
    CommunicationFault,
}

pub enum EnumScope {
    PrivateConnections,
    PublicConnections,
    AllConnections,
    Public,
    Private,
    All,
}
