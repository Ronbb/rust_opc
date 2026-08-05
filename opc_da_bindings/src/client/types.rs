use windows::Win32::Foundation::FILETIME;
use windows::Win32::System::Variant::VARIANT;
use windows_core::HRESULT;

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ServerItemHandle(pub(crate) u32);

impl ServerItemHandle {
    pub fn raw(self) -> u32 {
        self.0
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ClientItemHandle(pub u32);

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ItemError {
    pub code: HRESULT,
}

impl std::fmt::Display for ItemError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "OPC item operation failed: {:?}", self.code)
    }
}

impl std::error::Error for ItemError {}

#[derive(Clone, Debug)]
pub struct ItemSpec {
    pub item_id: String,
    pub access_path: String,
    pub active: bool,
    pub client_handle: ClientItemHandle,
    pub requested_data_type: u16,
    pub blob: Vec<u8>,
}

impl ItemSpec {
    pub fn new(item_id: impl Into<String>, client_handle: u32) -> Self {
        Self {
            item_id: item_id.into(),
            access_path: String::new(),
            active: true,
            client_handle: ClientItemHandle(client_handle),
            requested_data_type: 0,
            blob: Vec::new(),
        }
    }
}

#[derive(Clone, Debug)]
pub struct AddedItem {
    pub server_handle: ServerItemHandle,
    pub canonical_data_type: u16,
    pub access_rights: u32,
    pub blob: Vec<u8>,
}

#[derive(Clone, Debug)]
pub struct Sample {
    pub client_handle: ClientItemHandle,
    pub timestamp: FILETIME,
    pub quality: u16,
    pub value: VARIANT,
}

#[derive(Clone, Debug)]
pub struct WriteValue {
    pub server_handle: ServerItemHandle,
    pub value: VARIANT,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DataSource {
    Cache,
    Device,
}

impl DataSource {
    pub(crate) fn as_abi(self) -> crate::tagOPCDATASOURCE {
        match self {
            Self::Cache => crate::OPC_DS_CACHE,
            Self::Device => crate::OPC_DS_DEVICE,
        }
    }
}

#[derive(Clone, Debug)]
pub struct GroupOptions {
    pub name: String,
    pub active: bool,
    pub requested_update_rate: u32,
    pub client_handle: u32,
    pub time_bias: Option<i32>,
    pub percent_deadband: Option<f32>,
    pub locale: u32,
}

impl GroupOptions {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            active: true,
            requested_update_rate: 1_000,
            client_handle: 0,
            time_bias: None,
            percent_deadband: None,
            locale: 0,
        }
    }
}

#[derive(Clone, Debug)]
pub struct GroupState {
    pub update_rate: u32,
    pub active: bool,
    pub name: String,
    pub time_bias: i32,
    pub percent_deadband: f32,
    pub locale: u32,
    pub client_handle: u32,
    pub server_handle: u32,
}

#[derive(Clone, Debug)]
pub struct ServerStatus {
    pub start_time: FILETIME,
    pub current_time: FILETIME,
    pub last_update_time: FILETIME,
    pub state: crate::tagOPCSERVERSTATE,
    pub group_count: u32,
    pub bandwidth: u32,
    pub version: (u16, u16, u16),
    pub vendor_info: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PropertyDescription {
    pub id: u32,
    pub description: String,
    pub data_type: u16,
}

#[derive(Clone, Debug)]
pub struct PropertyValue {
    pub id: u32,
    pub value: VARIANT,
}
