use std::mem::ManuallyDrop;

use windows::Win32::Foundation::E_INVALIDARG;

use crate::com::{
    base::{SystemTime, Variant},
    bindings,
    utils::{PointerWriter, TryWriteArray, TryWriteTo},
};

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

pub struct FormatEtc {}

pub struct StorageMedium {}

pub struct DataSource {}

impl TryFrom<ItemProperties> for opc_da_bindings::tagOPCITEMPROPERTIES {
    type Error = windows_core::Error;

    fn try_from(value: ItemProperties) -> Result<Self, Self::Error> {
        let result = opc_da_bindings::tagOPCITEMPROPERTIES {
            hrErrorID: value.error_id,
            dwNumProperties: value.item_properties.len() as u32,
            pItemProperties: std::ptr::null_mut(),
            dwReserved: 0,
        };

        PointerWriter::try_write_array(
            &value
                .item_properties
                .into_iter()
                .map(|item_property| match item_property.try_into() {
                    Ok(item_property) => item_property,
                    Err(error) => {
                        let mut result = opc_da_bindings::tagOPCITEMPROPERTY::default();
                        result.hrErrorID = (error as windows_core::Error).code();
                        result
                    }
                })
                .collect::<Vec<_>>(),
            result.pItemProperties,
        )?;

        Ok(result)
    }
}

impl TryFrom<ItemProperty> for opc_da_bindings::tagOPCITEMPROPERTY {
    type Error = windows_core::Error;

    fn try_from(value: ItemProperty) -> Result<Self, Self::Error> {
        Ok(opc_da_bindings::tagOPCITEMPROPERTY {
            vtDataType: value.data_type,
            wReserved: 0,
            dwPropertyID: value.property_id,
            szItemID: PointerWriter::try_write_to(&value.item_id)?,
            szDescription: PointerWriter::try_write_to(&value.description)?,
            vValue: ManuallyDrop::new(value.value.into()),
            hrErrorID: value.error_id,
            dwReserved: 0,
        })
    }
}

impl From<BrowseFilter> for opc_da_bindings::tagOPCBROWSEFILTER {
    fn from(value: BrowseFilter) -> Self {
        match value {
            BrowseFilter::All => opc_da_bindings::OPC_BROWSE_FILTER_ALL,
            BrowseFilter::Branches => opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES,
            BrowseFilter::Items => opc_da_bindings::OPC_BROWSE_FILTER_ITEMS,
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCBROWSEFILTER> for BrowseFilter {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCBROWSEFILTER) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_BROWSE_FILTER_ALL => Ok(BrowseFilter::All),
            opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES => Ok(BrowseFilter::Branches),
            opc_da_bindings::OPC_BROWSE_FILTER_ITEMS => Ok(BrowseFilter::Items),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseFilter",
            )),
        }
    }
}

impl TryFrom<BrowseElement> for opc_da_bindings::tagOPCBROWSEELEMENT {
    type Error = windows_core::Error;

    fn try_from(value: BrowseElement) -> Result<Self, Self::Error> {
        Ok(opc_da_bindings::tagOPCBROWSEELEMENT {
            szName: PointerWriter::try_write_to(&value.name)?,
            szItemID: PointerWriter::try_write_to(&value.item_id)?,
            dwFlagValue: value.flag_value,
            dwReserved: 0,
            ItemProperties: value.item_properties.try_into()?,
        })
    }
}

impl From<NamespaceType> for opc_da_bindings::tagOPCNAMESPACETYPE {
    fn from(value: NamespaceType) -> Self {
        match value {
            NamespaceType::Flat => opc_da_bindings::OPC_NS_FLAT,
            NamespaceType::Hierarchical => opc_da_bindings::OPC_NS_HIERARCHIAL,
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCNAMESPACETYPE> for NamespaceType {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCNAMESPACETYPE) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_NS_FLAT => Ok(NamespaceType::Flat),
            opc_da_bindings::OPC_NS_HIERARCHIAL => Ok(NamespaceType::Hierarchical),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid NamespaceType",
            )),
        }
    }
}

impl TryFrom<(opc_da_bindings::tagOPCBROWSEDIRECTION, String)> for BrowseDirection {
    type Error = windows_core::Error;

    fn try_from(
        value: (opc_da_bindings::tagOPCBROWSEDIRECTION, String),
    ) -> Result<Self, Self::Error> {
        match value {
            (opc_da_bindings::OPC_BROWSE_UP, _) => Ok(BrowseDirection::Up),
            (opc_da_bindings::OPC_BROWSE_DOWN, _) => Ok(BrowseDirection::Down),
            (opc_da_bindings::OPC_BROWSE_TO, name) => Ok(BrowseDirection::To(name)),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseDirection",
            )),
        }
    }
}

impl From<BrowseDirection> for (opc_da_bindings::tagOPCBROWSEDIRECTION, String) {
    fn from(value: BrowseDirection) -> Self {
        match value {
            BrowseDirection::Up => (opc_da_bindings::OPC_BROWSE_UP, String::new()),
            BrowseDirection::Down => (opc_da_bindings::OPC_BROWSE_DOWN, String::new()),
            BrowseDirection::To(name) => (opc_da_bindings::OPC_BROWSE_TO, name),
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCBROWSETYPE> for BrowseType {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCBROWSETYPE) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_BRANCH => Ok(BrowseType::Branch),
            opc_da_bindings::OPC_LEAF => Ok(BrowseType::Leaf),
            opc_da_bindings::OPC_FLAT => Ok(BrowseType::Flat),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseFilter",
            )),
        }
    }
}

impl From<BrowseType> for opc_da_bindings::tagOPCBROWSETYPE {
    fn from(value: BrowseType) -> Self {
        match value {
            BrowseType::Branch => opc_da_bindings::OPC_BRANCH,
            BrowseType::Leaf => opc_da_bindings::OPC_LEAF,
            BrowseType::Flat => opc_da_bindings::OPC_FLAT,
        }
    }
}

impl TryFrom<bindings::tagOPCITEMVQT> for OptionalVqt {
    type Error = windows_core::Error;

    fn try_from(value: bindings::tagOPCITEMVQT) -> Result<Self, Self::Error> {
        Ok(OptionalVqt {
            value: value.vDataValue.as_raw().clone().into(),
            quality: if value.bQualitySpecified.as_bool() {
                Some(value.wQuality)
            } else {
                None
            },
            timestamp: if value.bTimeStampSpecified.as_bool() {
                Some(value.ftTimeStamp.into())
            } else {
                None
            },
        })
    }
}

impl TryFrom<ServerStatus> for bindings::tagOPCSERVERSTATUS {
    type Error = windows_core::Error;

    fn try_from(value: ServerStatus) -> Result<Self, Self::Error> {
        Ok(Self {
            ftStartTime: value.start_time.into(),
            ftCurrentTime: value.current_time.into(),
            ftLastUpdateTime: value.last_update_time.into(),
            dwServerState: value.server_state.into(),
            dwGroupCount: value.group_count,
            dwBandWidth: value.band_width,
            wMajorVersion: value.major_version,
            wMinorVersion: value.minor_version,
            wBuildNumber: value.build_number,
            szVendorInfo: PointerWriter::try_write_to(&value.vendor_info)?,
            wReserved: 0,
        })
    }
}

impl From<ServerState> for bindings::tagOPCSERVERSTATE {
    fn from(value: ServerState) -> Self {
        match value {
            ServerState::Running => bindings::OPC_STATUS_RUNNING,
            ServerState::Failed => bindings::OPC_STATUS_FAILED,
            ServerState::NoConfig => bindings::OPC_STATUS_NOCONFIG,
            ServerState::Suspended => bindings::OPC_STATUS_SUSPENDED,
            ServerState::Test => bindings::OPC_STATUS_TEST,
            ServerState::CommunicationFault => bindings::OPC_STATUS_COMM_FAULT,
        }
    }
}

impl TryFrom<bindings::tagOPCENUMSCOPE> for EnumScope {
    type Error = windows_core::Error;

    fn try_from(value: bindings::tagOPCENUMSCOPE) -> Result<Self, Self::Error> {
        match value {
            bindings::OPC_ENUM_PRIVATE_CONNECTIONS => Ok(EnumScope::PrivateConnections),
            bindings::OPC_ENUM_PUBLIC_CONNECTIONS => Ok(EnumScope::PublicConnections),
            bindings::OPC_ENUM_ALL_CONNECTIONS => Ok(EnumScope::AllConnections),
            bindings::OPC_ENUM_PUBLIC => Ok(EnumScope::Public),
            bindings::OPC_ENUM_PRIVATE => Ok(EnumScope::Private),
            bindings::OPC_ENUM_ALL => Ok(EnumScope::All),
            _ => Err(windows_core::Error::new(E_INVALIDARG, "Invalid EnumScope")),
        }
    }
}
