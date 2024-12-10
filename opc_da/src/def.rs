use crate::client::{LocalPointer, RemoteArray};

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

#[derive(Debug, Clone, Default, PartialEq)]
pub struct GroupState {
    pub update_rate: u32,
    pub active: bool,
    pub name: String,
    pub time_bias: i32,
    pub percent_deadband: f32,
    pub locale_id: u32,
    pub client_group_handle: u32,
    pub server_group_handle: u32,
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
            dwBlobSize: self.blob.len().try_into()?,
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
