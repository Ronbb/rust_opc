use std::mem::ManuallyDrop;

use crate::{
    try_from_native,
    utils::{IntoBridge, LocalPointer, RemoteArray, ToNative, TryFromNative, TryToNative},
};

#[derive(Debug, Clone, PartialEq)]
pub enum Version {
    V1,
    V2,
    V3,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct GroupState {
    pub update_rate: u32,
    pub active: bool,
    pub name: String,
    pub time_bias: i32,
    pub percent_deadband: f32,
    pub locale_id: u32,
    pub client_handle: u32,
    pub server_handle: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ServerStatus {
    pub start_time: std::time::SystemTime,
    pub current_time: std::time::SystemTime,
    pub last_update_time: std::time::SystemTime,
    pub server_state: ServerState,
    pub group_count: u32,
    pub band_width: u32,
    pub major_version: u16,
    pub minor_version: u16,
    pub build_number: u16,
    pub vendor_info: String,
}

impl TryFromNative<opc_da_bindings::tagOPCSERVERSTATUS> for ServerStatus {
    fn try_from_native(
        native: &opc_da_bindings::tagOPCSERVERSTATUS,
    ) -> windows::core::Result<Self> {
        Ok(Self {
            start_time: try_from_native!(&native.ftStartTime),
            current_time: try_from_native!(&native.ftCurrentTime),
            last_update_time: try_from_native!(&native.ftLastUpdateTime),
            server_state: try_from_native!(&native.dwServerState),
            group_count: native.dwGroupCount,
            band_width: native.dwBandWidth,
            major_version: native.wMajorVersion,
            minor_version: native.wMinorVersion,
            build_number: native.wBuildNumber,
            vendor_info: try_from_native!(&native.szVendorInfo),
        })
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct ItemDef {
    pub access_path: String,
    pub item_id: String,
    pub active: bool,
    pub client_handle: u32,
    pub data_type: u16,
    pub blob: Vec<u8>,
}

pub mod bridge {
    use crate::utils::LocalPointer;

    pub struct ItemDef {
        pub access_path: LocalPointer<Vec<u16>>,
        pub item_id: LocalPointer<Vec<u16>>,
        pub active: bool,
        pub item_client_handle: u32,
        pub requested_data_type: u16,
        pub blob: LocalPointer<Vec<u8>>,
    }
}

impl IntoBridge<bridge::ItemDef> for ItemDef {
    fn into_bridge(self) -> bridge::ItemDef {
        bridge::ItemDef {
            access_path: LocalPointer::from(&self.access_path),
            item_id: LocalPointer::from(&self.item_id),
            active: self.active,
            item_client_handle: self.client_handle,
            requested_data_type: self.data_type,
            blob: LocalPointer::new(Some(self.blob)),
        }
    }
}

impl TryToNative<opc_da_bindings::tagOPCITEMDEF> for bridge::ItemDef {
    fn try_to_native(&self) -> windows::core::Result<opc_da_bindings::tagOPCITEMDEF> {
        Ok(opc_da_bindings::tagOPCITEMDEF {
            szAccessPath: self.access_path.as_pwstr(),
            szItemID: self.item_id.as_pwstr(),
            bActive: self.active.into(),
            hClient: self.item_client_handle,
            vtRequestedDataType: self.requested_data_type,
            dwBlobSize: self.blob.len().try_into().map_err(|_| {
                windows::core::Error::new(
                    windows::Win32::Foundation::E_INVALIDARG,
                    "Blob size exceeds u32 maximum value",
                )
            })?,
            pBlob: self.blob.as_array_ptr() as *mut _,
            wReserved: 0,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ItemResult {
    pub server_handle: u32,
    pub data_type: u16,
    pub access_rights: u32,
    pub blob: Vec<u8>,
}

impl TryFromNative<opc_da_bindings::tagOPCITEMRESULT> for ItemResult {
    fn try_from_native(native: &opc_da_bindings::tagOPCITEMRESULT) -> windows::core::Result<Self> {
        Ok(Self {
            server_handle: native.hServer,
            data_type: native.vtCanonicalDataType,
            access_rights: native.dwAccessRights,
            blob: RemoteArray::from_mut_ptr(native.pBlob, native.dwBlobSize)
                .as_slice()
                .to_vec(),
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ServerState {
    Running,
    Failed,
    NoConfig,
    Suspended,
    Test,
    CommunicationFault,
}

impl TryFromNative<opc_da_bindings::tagOPCSERVERSTATE> for ServerState {
    fn try_from_native(native: &opc_da_bindings::tagOPCSERVERSTATE) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_STATUS_RUNNING => Ok(ServerState::Running),
            opc_da_bindings::OPC_STATUS_FAILED => Ok(ServerState::Failed),
            opc_da_bindings::OPC_STATUS_NOCONFIG => Ok(ServerState::NoConfig),
            opc_da_bindings::OPC_STATUS_SUSPENDED => Ok(ServerState::Suspended),
            opc_da_bindings::OPC_STATUS_TEST => Ok(ServerState::Test),
            opc_da_bindings::OPC_STATUS_COMM_FAULT => Ok(ServerState::CommunicationFault),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown server state: {:?}", unknown),
            )),
        }
    }
}

impl ToNative<opc_da_bindings::tagOPCSERVERSTATE> for ServerState {
    fn to_native(&self) -> opc_da_bindings::tagOPCSERVERSTATE {
        match self {
            ServerState::Running => opc_da_bindings::OPC_STATUS_RUNNING,
            ServerState::Failed => opc_da_bindings::OPC_STATUS_FAILED,
            ServerState::NoConfig => opc_da_bindings::OPC_STATUS_NOCONFIG,
            ServerState::Suspended => opc_da_bindings::OPC_STATUS_SUSPENDED,
            ServerState::Test => opc_da_bindings::OPC_STATUS_TEST,
            ServerState::CommunicationFault => opc_da_bindings::OPC_STATUS_COMM_FAULT,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum EnumScope {
    PrivateConnections,
    PublicConnections,
    AllConnections,
    Public,
    Private,
    All,
}

impl TryFromNative<opc_da_bindings::tagOPCENUMSCOPE> for EnumScope {
    fn try_from_native(native: &opc_da_bindings::tagOPCENUMSCOPE) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_ENUM_PRIVATE_CONNECTIONS => Ok(EnumScope::PrivateConnections),
            opc_da_bindings::OPC_ENUM_PUBLIC_CONNECTIONS => Ok(EnumScope::PublicConnections),
            opc_da_bindings::OPC_ENUM_ALL_CONNECTIONS => Ok(EnumScope::AllConnections),
            opc_da_bindings::OPC_ENUM_PUBLIC => Ok(EnumScope::Public),
            opc_da_bindings::OPC_ENUM_PRIVATE => Ok(EnumScope::Private),
            opc_da_bindings::OPC_ENUM_ALL => Ok(EnumScope::All),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown enum scope: {:?}", unknown),
            )),
        }
    }
}

impl ToNative<opc_da_bindings::tagOPCENUMSCOPE> for EnumScope {
    fn to_native(&self) -> opc_da_bindings::tagOPCENUMSCOPE {
        match self {
            EnumScope::PrivateConnections => opc_da_bindings::OPC_ENUM_PRIVATE_CONNECTIONS,
            EnumScope::PublicConnections => opc_da_bindings::OPC_ENUM_PUBLIC_CONNECTIONS,
            EnumScope::AllConnections => opc_da_bindings::OPC_ENUM_ALL_CONNECTIONS,
            EnumScope::Public => opc_da_bindings::OPC_ENUM_PUBLIC,
            EnumScope::Private => opc_da_bindings::OPC_ENUM_PRIVATE,
            EnumScope::All => opc_da_bindings::OPC_ENUM_ALL,
        }
    }
}

pub struct ItemAttributes {
    pub access_path: String,
    pub item_id: String,
    pub active: bool,
    pub client_handle: u32,
    pub server_handle: u32,
    pub access_rights: u32,
    pub blob: Vec<u8>,
    pub requested_data_type: u16,
    pub canonical_data_type: u16,
    pub eu_type: EuType,
    pub eu_info: windows_core::VARIANT,
}

impl TryFromNative<opc_da_bindings::tagOPCITEMATTRIBUTES> for ItemAttributes {
    fn try_from_native(
        native: &opc_da_bindings::tagOPCITEMATTRIBUTES,
    ) -> windows::core::Result<Self> {
        Ok(Self {
            access_path: try_from_native!(&native.szAccessPath),
            item_id: try_from_native!(&native.szItemID),
            active: native.bActive.into(),
            client_handle: native.hClient,
            server_handle: native.hServer,
            access_rights: native.dwAccessRights,
            blob: RemoteArray::from_mut_ptr(native.pBlob, native.dwBlobSize)
                .as_slice()
                .to_vec(),
            requested_data_type: native.vtRequestedDataType,
            canonical_data_type: native.vtCanonicalDataType,
            eu_type: try_from_native!(&native.dwEUType),
            eu_info: (*native.vEUInfo).clone(),
        })
    }
}

pub enum EuType {
    NoEnum,
    Analog,
    Enumerated,
}

impl TryFromNative<opc_da_bindings::tagOPCEUTYPE> for EuType {
    fn try_from_native(native: &opc_da_bindings::tagOPCEUTYPE) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_NOENUM => Ok(EuType::NoEnum),
            opc_da_bindings::OPC_ANALOG => Ok(EuType::Analog),
            opc_da_bindings::OPC_ENUMERATED => Ok(EuType::Enumerated),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown EU type: {:?}", unknown),
            )),
        }
    }
}

pub struct ItemState {
    pub client_handle: u32,
    pub timestamp: std::time::SystemTime,
    pub quality: u16,
    pub data_value: windows::Win32::System::Variant::VARIANT,
}

impl TryFromNative<opc_da_bindings::tagOPCITEMSTATE> for ItemState {
    fn try_from_native(native: &opc_da_bindings::tagOPCITEMSTATE) -> windows::core::Result<Self> {
        Ok(Self {
            client_handle: native.hClient,
            timestamp: try_from_native!(&native.ftTimeStamp),
            quality: native.wQuality,
            data_value: (*native.vDataValue).clone(),
        })
    }
}

pub enum DataSourceTarget {
    ForceCache,
    ForceDevice,
    WithMaxAge(u32),
}

impl DataSourceTarget {
    pub fn max_age(&self) -> u32 {
        match self {
            DataSourceTarget::WithMaxAge(max_age) => *max_age,
            DataSourceTarget::ForceCache => u32::MAX,
            DataSourceTarget::ForceDevice => 0,
        }
    }
}

impl TryFromNative<opc_da_bindings::tagOPCDATASOURCE> for DataSourceTarget {
    fn try_from_native(native: &opc_da_bindings::tagOPCDATASOURCE) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_DS_CACHE => Ok(DataSourceTarget::ForceCache),
            opc_da_bindings::OPC_DS_DEVICE => Ok(DataSourceTarget::ForceDevice),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown data source: {:?}", unknown),
            )),
        }
    }
}

impl TryToNative<opc_da_bindings::tagOPCDATASOURCE> for DataSourceTarget {
    fn try_to_native(&self) -> windows::core::Result<opc_da_bindings::tagOPCDATASOURCE> {
        match self {
            DataSourceTarget::ForceCache => Ok(opc_da_bindings::OPC_DS_CACHE),
            DataSourceTarget::ForceDevice => Ok(opc_da_bindings::OPC_DS_DEVICE),
            DataSourceTarget::WithMaxAge(_) => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "MaxAge data source requires a value",
            )),
        }
    }
}

pub struct ItemValue {
    pub value: windows::Win32::System::Variant::VARIANT,
    pub quality: u16,
    pub timestamp: std::time::SystemTime,
}

impl
    TryFromNative<(
        RemoteArray<windows::Win32::System::Variant::VARIANT>,
        RemoteArray<u16>,
        RemoteArray<windows::Win32::Foundation::FILETIME>,
        RemoteArray<windows::core::HRESULT>,
    )> for Vec<windows::core::Result<ItemValue>>
{
    fn try_from_native(
        native: &(
            RemoteArray<windows::Win32::System::Variant::VARIANT>,
            RemoteArray<u16>,
            RemoteArray<windows::Win32::Foundation::FILETIME>,
            RemoteArray<windows::core::HRESULT>,
        ),
    ) -> windows::core::Result<Self> {
        let (values, qualities, timestamps, errors) = native;

        if values.len() != qualities.len()
            || values.len() != timestamps.len()
            || values.len() != errors.len()
        {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Arrays have different lengths",
            ));
        }

        Ok(values
            .as_slice()
            .iter()
            .zip(qualities.as_slice())
            .zip(timestamps.as_slice())
            .zip(errors.as_slice())
            .map(|(((value, quality), timestamp), error)| {
                if error.is_ok() {
                    Ok(ItemValue {
                        value: value.clone(),
                        quality: *quality,
                        timestamp: try_from_native!(timestamp),
                    })
                } else {
                    Err((*error).into())
                }
            })
            .collect())
    }
}

pub struct ItemPartialValue {
    pub value: windows::Win32::System::Variant::VARIANT,
    pub quality: Option<u16>,
    pub timestamp: Option<std::time::SystemTime>,
}

// try to native
impl TryToNative<opc_da_bindings::tagOPCITEMVQT> for ItemPartialValue {
    fn try_to_native(&self) -> windows::core::Result<opc_da_bindings::tagOPCITEMVQT> {
        Ok(opc_da_bindings::tagOPCITEMVQT {
            vDataValue: ManuallyDrop::new(self.value.clone()),
            bQualitySpecified: self.quality.is_some().into(),
            wQuality: self.quality.unwrap_or_default(),
            bTimeStampSpecified: self.timestamp.is_some().into(),
            ftTimeStamp: self
                .timestamp
                .map(|t| t.try_to_native())
                .transpose()?
                .unwrap_or_default(),
            wReserved: 0,
            dwReserved: 0,
        })
    }
}

pub enum BrowseType {
    Branch,
    Leaf,
    Flat,
}

impl TryFromNative<opc_da_bindings::tagOPCBROWSETYPE> for BrowseType {
    fn try_from_native(native: &opc_da_bindings::tagOPCBROWSETYPE) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_BRANCH => Ok(BrowseType::Branch),
            opc_da_bindings::OPC_LEAF => Ok(BrowseType::Leaf),
            opc_da_bindings::OPC_FLAT => Ok(BrowseType::Flat),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown browse type: {:?}", unknown),
            )),
        }
    }
}

impl ToNative<opc_da_bindings::tagOPCBROWSETYPE> for BrowseType {
    fn to_native(&self) -> opc_da_bindings::tagOPCBROWSETYPE {
        match self {
            BrowseType::Branch => opc_da_bindings::OPC_BRANCH,
            BrowseType::Leaf => opc_da_bindings::OPC_LEAF,
            BrowseType::Flat => opc_da_bindings::OPC_FLAT,
        }
    }
}

pub enum BrowseFilter {
    All,
    Branches,
    Items,
}

impl TryFromNative<opc_da_bindings::tagOPCBROWSEFILTER> for BrowseFilter {
    fn try_from_native(
        native: &opc_da_bindings::tagOPCBROWSEFILTER,
    ) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_BROWSE_FILTER_ALL => Ok(BrowseFilter::All),
            opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES => Ok(BrowseFilter::Branches),
            opc_da_bindings::OPC_BROWSE_FILTER_ITEMS => Ok(BrowseFilter::Items),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown browse filter: {:?}", unknown),
            )),
        }
    }
}

impl ToNative<opc_da_bindings::tagOPCBROWSEFILTER> for BrowseFilter {
    fn to_native(&self) -> opc_da_bindings::tagOPCBROWSEFILTER {
        match self {
            BrowseFilter::All => opc_da_bindings::OPC_BROWSE_FILTER_ALL,
            BrowseFilter::Branches => opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES,
            BrowseFilter::Items => opc_da_bindings::OPC_BROWSE_FILTER_ITEMS,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DataCallbackEvent {
    DataChange(DataChangeEvent),
    ReadComplete(ReadCompleteEvent),
    WriteComplete(WriteCompleteEvent),
    CancelComplete(CancelCompleteEvent),
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataChangeEvent {
    pub transaction_id: u32,
    pub group_handle: u32,
    pub master_quality: windows_core::HRESULT,
    pub master_error: windows_core::HRESULT,
    pub client_items: RemoteArray<u32>,
    pub values: RemoteArray<windows::Win32::System::Variant::VARIANT>,
    pub qualities: RemoteArray<u16>,
    pub timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
    pub errors: RemoteArray<windows_core::HRESULT>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ReadCompleteEvent {
    pub transaction_id: u32,
    pub group_handle: u32,
    pub master_quality: windows_core::HRESULT,
    pub master_error: windows_core::HRESULT,
    pub client_items: RemoteArray<u32>,
    pub values: RemoteArray<windows::Win32::System::Variant::VARIANT>,
    pub qualities: RemoteArray<u16>,
    pub timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
    pub errors: RemoteArray<windows_core::HRESULT>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct WriteCompleteEvent {
    pub transaction_id: u32,
    pub group_handle: u32,
    pub master_error: windows_core::HRESULT,
    pub client_handles: RemoteArray<u32>,
    pub errors: RemoteArray<windows_core::HRESULT>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CancelCompleteEvent {
    pub transaction_id: u32,
    pub group_handle: u32,
}

pub enum NamespaceType {
    Flat,
    Hierarchy,
}

impl TryFromNative<opc_da_bindings::tagOPCNAMESPACETYPE> for NamespaceType {
    fn try_from_native(
        native: &opc_da_bindings::tagOPCNAMESPACETYPE,
    ) -> windows::core::Result<Self> {
        match *native {
            opc_da_bindings::OPC_NS_HIERARCHIAL => Ok(NamespaceType::Hierarchy),
            opc_da_bindings::OPC_NS_FLAT => Ok(NamespaceType::Flat),
            unknown => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                format!("Unknown namespace type: {:?}", unknown),
            )),
        }
    }
}

impl ToNative<opc_da_bindings::tagOPCNAMESPACETYPE> for NamespaceType {
    fn to_native(&self) -> opc_da_bindings::tagOPCNAMESPACETYPE {
        match self {
            NamespaceType::Hierarchy => opc_da_bindings::OPC_NS_HIERARCHIAL,
            NamespaceType::Flat => opc_da_bindings::OPC_NS_FLAT,
        }
    }
}
