use windows_core::Interface as _;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ServerVersion {
    Version10,
    Version20,
    Version30,
}

impl ServerVersion {
    pub fn to_guid(&self) -> windows_core::GUID {
        match self {
            ServerVersion::Version10 => opc_da_bindings::CATID_OPCDAServer10::IID,
            ServerVersion::Version20 => opc_da_bindings::CATID_OPCDAServer20::IID,
            ServerVersion::Version30 => opc_da_bindings::CATID_OPCDAServer30::IID,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ServerFilter {
    pub(super) available_versions: Vec<ServerVersion>,
    pub(super) requires_versions: Vec<ServerVersion>,
}

impl Default for ServerFilter {
    fn default() -> Self {
        Self {
            available_versions: vec![
                ServerVersion::Version10,
                ServerVersion::Version20,
                ServerVersion::Version30,
            ],
            requires_versions: vec![
                ServerVersion::Version10,
                ServerVersion::Version20,
                ServerVersion::Version30,
            ],
        }
    }
}

impl ServerFilter {
    pub fn with_version(mut self, version: ServerVersion) -> Self {
        self.available_versions = vec![version];

        self
    }

    pub fn with_versions<I: IntoIterator<Item = ServerVersion>>(mut self, versions: I) -> Self {
        let versions = versions
            .into_iter()
            .collect::<std::collections::HashSet<_>>();

        self.available_versions = versions.into_iter().collect();

        self
    }

    pub fn with_all_versions(mut self) -> Self {
        self.available_versions = vec![
            ServerVersion::Version10,
            ServerVersion::Version20,
            ServerVersion::Version30,
        ];

        self
    }
}
