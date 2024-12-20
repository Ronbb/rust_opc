use crate::client::{LocalPointer, RemoteArray, RemotePointer};

pub(crate) trait IntoBridge<Bridge> {
    fn into_bridge(self) -> Bridge;
}

pub(crate) trait ToNative<Native> {
    fn to_native(&self) -> Native;
}

pub(crate) trait FromNative<Native> {
    fn from_native(native: &Native) -> Self
    where
        Self: Sized;
}

pub(crate) trait TryToNative<Native> {
    fn try_to_native(&self) -> windows::core::Result<Native>;
}

pub(crate) trait TryFromNative<Native> {
    fn try_from_native(native: &Native) -> windows::core::Result<Self>
    where
        Self: Sized;
}

pub(crate) trait TryToLocal<Local> {
    fn try_to_local(&self) -> windows::core::Result<Local>;
}

impl<Native, T: TryFromNative<Native>> TryToLocal<T> for Native {
    fn try_to_local(&self) -> windows::core::Result<T> {
        T::try_from_native(self)
    }
}

impl<Native, T: FromNative<Native>> TryFromNative<Native> for T {
    fn try_from_native(native: &Native) -> windows::core::Result<Self> {
        Ok(Self::from_native(native))
    }
}

impl<Native, T: ToNative<Native>> TryToNative<Native> for T {
    fn try_to_native(&self) -> windows::core::Result<Native> {
        Ok(self.to_native())
    }
}

impl<Bridge, B: IntoBridge<Bridge>> IntoBridge<Vec<Bridge>> for Vec<B> {
    fn into_bridge(self) -> Vec<Bridge> {
        self.into_iter().map(IntoBridge::into_bridge).collect()
    }
}

impl<Bridge, B: IntoBridge<Bridge> + Clone> IntoBridge<Vec<Bridge>> for &[B] {
    fn into_bridge(self) -> Vec<Bridge> {
        self.iter().cloned().map(IntoBridge::into_bridge).collect()
    }
}

impl<Native, T: TryToNative<Native>> TryToNative<Vec<Native>> for Vec<T> {
    fn try_to_native(&self) -> windows::core::Result<Vec<Native>> {
        self.iter().map(TryToNative::try_to_native).collect()
    }
}

impl TryFromNative<RemoteArray<windows::core::HRESULT>> for Vec<windows::core::Result<()>> {
    fn try_from_native(
        native: &RemoteArray<windows::core::HRESULT>,
    ) -> windows::core::Result<Self> {
        Ok(native.as_slice().iter().map(|v| (*v).ok()).collect())
    }
}

impl<Native, T: TryFromNative<Native>> TryFromNative<RemoteArray<Native>> for Vec<T> {
    fn try_from_native(native: &RemoteArray<Native>) -> windows::core::Result<Self> {
        native.as_slice().iter().map(T::try_from_native).collect()
    }
}
impl<Native, T: TryFromNative<Native>>
    TryFromNative<(RemoteArray<Native>, RemoteArray<windows::core::HRESULT>)>
    for Vec<windows::core::Result<T>>
{
    fn try_from_native(
        native: &(RemoteArray<Native>, RemoteArray<windows::core::HRESULT>),
    ) -> windows::core::Result<Self> {
        let (results, errors) = native;
        if results.len() != errors.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Results and errors arrays have different lengths",
            ));
        }

        Ok(results
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(result, error)| {
                if error.is_ok() {
                    T::try_from_native(result)
                } else {
                    Err((*error).into())
                }
            })
            .collect())
    }
}

impl TryFromNative<windows::Win32::Foundation::FILETIME> for std::time::SystemTime {
    fn try_from_native(
        native: &windows::Win32::Foundation::FILETIME,
    ) -> windows::core::Result<Self> {
        let ft = ((native.dwHighDateTime as u64) << 32) | (native.dwLowDateTime as u64);
        let duration_since_1601 = std::time::Duration::from_nanos(ft * 100);

        let windows_to_unix_epoch_diff = std::time::Duration::from_secs(11_644_473_600);
        let duration_since_unix_epoch = duration_since_1601
            .checked_sub(windows_to_unix_epoch_diff)
            .ok_or_else(|| {
                windows::core::Error::new(
                    windows::Win32::Foundation::E_INVALIDARG,
                    "FILETIME is before UNIX_EPOCH",
                )
            })?;

        Ok(std::time::UNIX_EPOCH + duration_since_unix_epoch)
    }
}

macro_rules! from {
    ($native:expr) => {
        TryFromNative::try_from_native($native)?
    };
}

impl TryToNative<windows::Win32::Foundation::FILETIME> for std::time::SystemTime {
    fn try_to_native(&self) -> windows::core::Result<windows::Win32::Foundation::FILETIME> {
        let duration_since_unix_epoch =
            self.duration_since(std::time::UNIX_EPOCH).map_err(|_| {
                windows::core::Error::new(
                    windows::Win32::Foundation::E_INVALIDARG,
                    "SystemTime is before UNIX_EPOCH",
                )
            })?;

        let duration_since_windows_epoch =
            duration_since_unix_epoch + std::time::Duration::from_secs(11_644_473_600);

        let ft = duration_since_windows_epoch.as_nanos() / 100;

        Ok(windows::Win32::Foundation::FILETIME {
            dwLowDateTime: ft as u32,
            dwHighDateTime: (ft >> 32) as u32,
        })
    }
}

impl TryFromNative<windows::core::PWSTR> for String {
    fn try_from_native(native: &windows::core::PWSTR) -> windows::core::Result<Self> {
        RemotePointer::from(*native).try_into()
    }
}

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
            start_time: from!(&native.ftStartTime),
            current_time: from!(&native.ftCurrentTime),
            last_update_time: from!(&native.ftLastUpdateTime),
            server_state: from!(&native.dwServerState),
            group_count: native.dwGroupCount,
            band_width: native.dwBandWidth,
            major_version: native.wMajorVersion,
            minor_version: native.wMinorVersion,
            build_number: native.wBuildNumber,
            vendor_info: from!(&native.szVendorInfo),
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
    use crate::client::LocalPointer;

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
            access_path: from!(&native.szAccessPath),
            item_id: from!(&native.szItemID),
            active: native.bActive.into(),
            client_handle: native.hClient,
            server_handle: native.hServer,
            access_rights: native.dwAccessRights,
            blob: RemoteArray::from_mut_ptr(native.pBlob, native.dwBlobSize)
                .as_slice()
                .to_vec(),
            requested_data_type: native.vtRequestedDataType,
            canonical_data_type: native.vtCanonicalDataType,
            eu_type: from!(&native.dwEUType),
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
    pub data_value: windows::core::VARIANT,
}

impl TryFromNative<opc_da_bindings::tagOPCITEMSTATE> for ItemState {
    fn try_from_native(native: &opc_da_bindings::tagOPCITEMSTATE) -> windows::core::Result<Self> {
        Ok(Self {
            client_handle: native.hClient,
            timestamp: from!(&native.ftTimeStamp),
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
    pub value: windows::core::VARIANT,
    pub quality: u16,
    pub timestamp: std::time::SystemTime,
}

impl
    TryFromNative<(
        RemoteArray<windows::core::VARIANT>,
        RemoteArray<u16>,
        RemoteArray<windows::Win32::Foundation::FILETIME>,
        RemoteArray<windows::core::HRESULT>,
    )> for Vec<windows::core::Result<ItemValue>>
{
    fn try_from_native(
        native: &(
            RemoteArray<windows::core::VARIANT>,
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
                        timestamp: from!(timestamp),
                    })
                } else {
                    Err((*error).into())
                }
            })
            .collect())
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
