use windows_core::Interface as _;

use crate::client::traits::ServerTrait;

use super::Group;

/*
    opc_da_bindings::IOPCServer,
    opc_da_bindings::IOPCCommon,
    windows::Win32::System::Com::IConnectionPointContainer,
    opc_da_bindings::IOPCItemProperties,
    opc_da_bindings::IOPCBrowse,
    opc_da_bindings::IOPCServerPublicGroups,
    opc_da_bindings::IOPCBrowseServerAddressSpace,
    opc_da_bindings::IOPCItemIO
*/
pub struct Server {
    // 1.0 required
    // 2.0 required
    // 3.0 required
    pub(crate) server: opc_da_bindings::IOPCServer,
    // 1.0 N/A
    // 2.0 required
    // 3.0 required
    pub(crate) common: Option<opc_da_bindings::IOPCCommon>,
    // 1.0 N/A
    // 2.0 required
    // 3.0 required
    pub(crate) connection_point_container:
        Option<windows::Win32::System::Com::IConnectionPointContainer>,
    // 1.0 N/A
    // 2.0 required
    // 3.0 N/A
    pub(crate) item_properties: Option<opc_da_bindings::IOPCItemProperties>,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) browse: Option<opc_da_bindings::IOPCBrowse>,
    // 1.0 optional
    // 2.0 optional
    // 3.0 N/A
    pub(crate) server_public_groups: Option<opc_da_bindings::IOPCServerPublicGroups>,
    // 1.0 optional
    // 2.0 optional
    // 3.0 N/A
    pub(crate) browse_server_address_space: Option<opc_da_bindings::IOPCBrowseServerAddressSpace>,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) item_io: Option<opc_da_bindings::IOPCItemIO>,
}

impl TryFrom<windows_core::IUnknown> for Server {
    type Error = windows_core::Error;

    fn try_from(value: windows_core::IUnknown) -> windows_core::Result<Self> {
        let server = value.cast()?;
        let common = value.cast().ok();
        let connection_point_container = value.cast().ok();
        let item_properties = value.cast().ok();
        let browse = value.cast().ok();
        let server_public_groups = value.cast().ok();
        let browse_server_address_space = value.cast().ok();
        let item_io = value.cast().ok();

        Ok(Self {
            server,
            common,
            connection_point_container,
            item_properties,
            browse,
            server_public_groups,
            browse_server_address_space,
            item_io,
        })
    }
}

impl ServerTrait<Group> for Server {
    fn server(&self) -> &opc_da_bindings::IOPCServer {
        &self.server
    }
}
