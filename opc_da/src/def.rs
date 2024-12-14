use windows_core::PWSTR;

use crate::client::{LocalPointer, RemoteArray, RemotePointer};

pub(crate) trait IntoBridge {
    type Bridge;

    fn into_bridge(self) -> Self::Bridge;
}

pub(crate) trait ToNative {
    type Native;

    fn to_native(&self) -> Self::Native;
}

pub(crate) trait FromNative {
    type Native;

    fn from_native(native: &Self::Native) -> Self
    where
        Self: Sized;
}

pub(crate) trait TryToNative {
    type Native;

    fn try_to_native(&self) -> windows::core::Result<Self::Native>;
}

pub(crate) trait TryFromNative {
    type Native;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self>
    where
        Self: Sized;
}

impl<T: FromNative> TryFromNative for T {
    type Native = T::Native;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
        Ok(Self::from_native(native))
    }
}

impl<T: ToNative> TryToNative for T {
    type Native = T::Native;

    fn try_to_native(&self) -> windows::core::Result<Self::Native> {
        Ok(self.to_native())
    }
}

impl<B: IntoBridge> IntoBridge for Vec<B> {
    type Bridge = Vec<B::Bridge>;

    fn into_bridge(self) -> Self::Bridge {
        self.into_iter().map(IntoBridge::into_bridge).collect()
    }
}

impl<B: IntoBridge + Clone> IntoBridge for &[B] {
    type Bridge = Vec<B::Bridge>;

    fn into_bridge(self) -> Self::Bridge {
        self.iter().cloned().map(IntoBridge::into_bridge).collect()
    }
}

impl<N: TryToNative> TryToNative for Vec<N> {
    type Native = Vec<N::Native>;

    fn try_to_native(&self) -> windows::core::Result<Self::Native> {
        self.iter().map(TryToNative::try_to_native).collect()
    }
}

impl TryFromNative for std::time::SystemTime {
    type Native = windows::Win32::Foundation::FILETIME;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
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

impl TryToNative for std::time::SystemTime {
    type Native = windows::Win32::Foundation::FILETIME;

    fn try_to_native(&self) -> windows::core::Result<Self::Native> {
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

        Ok(Self::Native {
            dwLowDateTime: ft as u32,
            dwHighDateTime: (ft >> 32) as u32,
        })
    }
}

impl TryFromNative for String {
    type Native = PWSTR;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
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

impl TryFromNative for ServerStatus {
    type Native = opc_da_bindings::tagOPCSERVERSTATUS;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
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
    pub item_client_handle: u32,
    pub requested_data_type: u16,
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

impl IntoBridge for ItemDef {
    type Bridge = bridge::ItemDef;

    fn into_bridge(self) -> Self::Bridge {
        Self::Bridge {
            access_path: LocalPointer::from(&self.access_path),
            item_id: LocalPointer::from(&self.item_id),
            active: self.active,
            item_client_handle: self.item_client_handle,
            requested_data_type: self.requested_data_type,
            blob: LocalPointer::new(Some(self.blob)),
        }
    }
}

impl TryToNative for bridge::ItemDef {
    type Native = opc_da_bindings::tagOPCITEMDEF;

    fn try_to_native(&self) -> windows::core::Result<Self::Native> {
        Ok(Self::Native {
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
    pub item_server_handle: u32,
    pub canonical_data_type: u16,
    pub access_rights: u32,
    pub blob_size: u32,
    pub blob: RemoteArray<u8>,
}

impl TryFromNative for ItemResult {
    type Native = opc_da_bindings::tagOPCITEMRESULT;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
        Ok(Self {
            item_server_handle: native.hServer,
            canonical_data_type: native.vtCanonicalDataType,
            access_rights: native.dwAccessRights,
            blob_size: native.dwBlobSize,
            blob: RemoteArray::from_raw(native.pBlob, native.dwBlobSize),
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

impl TryFromNative for ServerState {
    type Native = opc_da_bindings::tagOPCSERVERSTATE;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
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

impl ToNative for ServerState {
    type Native = opc_da_bindings::tagOPCSERVERSTATE;

    fn to_native(&self) -> Self::Native {
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

impl TryFromNative for EnumScope {
    type Native = opc_da_bindings::tagOPCENUMSCOPE;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
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

impl ToNative for EnumScope {
    type Native = opc_da_bindings::tagOPCENUMSCOPE;

    fn to_native(&self) -> Self::Native {
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
