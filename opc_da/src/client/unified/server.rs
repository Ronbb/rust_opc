use windows::core::Interface as _;

use crate::client::{
    traits::{
        BrowseServerAddressSpaceTrait, ItemPropertiesTrait, ServerPublicGroupsTrait, ServerTrait,
    },
    BrowseTrait, CommonTrait, ConnectionPointContainerTrait, ItemIoTrait,
};

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

impl TryFrom<windows::core::IUnknown> for Server {
    type Error = windows::core::Error;

    fn try_from(value: windows::core::IUnknown) -> windows::core::Result<Self> {
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
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCServer> {
        Ok(&self.server)
    }
}

impl CommonTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCCommon> {
        self.common.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCCommon not supported",
            )
        })
    }
}

impl ConnectionPointContainerTrait for Server {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        self.connection_point_container.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IConnectionPointContainer not supported",
            )
        })
    }
}

impl ServerPublicGroupsTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCServerPublicGroups> {
        self.server_public_groups.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCServerPublicGroups not supported",
            )
        })
    }
}

impl BrowseServerAddressSpaceTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCBrowseServerAddressSpace> {
        self.browse_server_address_space.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCBrowseServerAddressSpace not supported",
            )
        })
    }
}

impl ItemPropertiesTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemProperties> {
        self.item_properties.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCItemProperties not supported",
            )
        })
    }
}

impl BrowseTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCBrowse> {
        self.browse.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCBrowse not supported",
            )
        })
    }
}

impl ItemIoTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemIO> {
        self.item_io.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCItemIO not supported",
            )
        })
    }
}
