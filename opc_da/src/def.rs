use windows_core::PWSTR;

use crate::client::{LocalPointer, RemoteArray, RemotePointer};

pub(crate) trait IntoBridge {
    type Bridge;

    fn into_bridge(self) -> Self::Bridge;
}

pub(crate) trait ToNative {
    type Native;

    fn to_native(&mut self) -> windows::core::Result<Self::Native>;
}

pub(crate) trait TryFromNative {
    type Native;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self>
    where
        Self: Sized;
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

impl<N: ToNative> ToNative for Vec<N> {
    type Native = Vec<N::Native>;

    fn to_native(&mut self) -> windows::core::Result<Self::Native> {
        self.iter_mut().map(ToNative::to_native).collect()
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

impl ToNative for std::time::SystemTime {
    type Native = windows::Win32::Foundation::FILETIME;

    fn to_native(&mut self) -> windows::core::Result<Self::Native> {
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
    pub server_state: i32,
    pub group_count: u32,
    pub band_width: u32,
    pub major_version: u16,
    pub minor_version: u16,
    pub build_number: u16,
    pub reserved: u16,
    pub vendor_info: String,
}

impl TryFromNative for ServerStatus {
    type Native = opc_da_bindings::tagOPCSERVERSTATUS;

    fn try_from_native(native: &Self::Native) -> windows::core::Result<Self> {
        Ok(Self {
            start_time: from!(&native.ftStartTime),
            current_time: from!(&native.ftCurrentTime),
            last_update_time: from!(&native.ftLastUpdateTime),
            server_state: native.dwServerState.0,
            group_count: native.dwGroupCount,
            band_width: native.dwBandWidth,
            major_version: native.wMajorVersion,
            minor_version: native.wMinorVersion,
            build_number: native.wBuildNumber,
            reserved: native.wReserved,
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

impl ToNative for bridge::ItemDef {
    type Native = opc_da_bindings::tagOPCITEMDEF;

    fn to_native(&mut self) -> windows::core::Result<Self::Native> {
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
            pBlob: self.blob.as_mut_array_ptr(),
            wReserved: 0,
        })
    }
}

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
