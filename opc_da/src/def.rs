use windows::core::Interface as _;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ServerVersion {
    Version10,
    Version20,
    Version30,
}

impl ServerVersion {
    const ALL: [ServerVersion; 3] = [
        ServerVersion::Version10,
        ServerVersion::Version20,
        ServerVersion::Version30,
    ];

    pub fn to_guid(&self) -> windows::core::GUID {
        match self {
            ServerVersion::Version10 => opc_da_bindings::CATID_OPCDAServer10::IID,
            ServerVersion::Version20 => opc_da_bindings::CATID_OPCDAServer20::IID,
            ServerVersion::Version30 => opc_da_bindings::CATID_OPCDAServer30::IID,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ServerFilter {
    pub(super) available_versions: std::collections::HashSet<ServerVersion>,
    pub(super) requires_versions: std::collections::HashSet<ServerVersion>,
}

impl Default for ServerFilter {
    fn default() -> Self {
        Self {
            available_versions: ServerVersion::ALL.into_iter().collect(),
            requires_versions: ServerVersion::ALL.into_iter().collect(),
        }
    }
}

impl ServerFilter {
    pub fn with_version(mut self, version: ServerVersion) -> Self {
        self.available_versions = std::iter::once(version).collect();
        self.requires_versions.retain(|v| *v == version);

        self
    }

    pub fn with_versions<I: IntoIterator<Item = ServerVersion>>(mut self, versions: I) -> Self {
        self.available_versions = versions.into_iter().collect();
        self.requires_versions
            .retain(|v| self.available_versions.contains(v));

        self
    }

    pub fn with_all_versions(mut self) -> Self {
        self.available_versions = ServerVersion::ALL.into_iter().collect();
        self
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
