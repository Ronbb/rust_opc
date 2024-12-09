use windows_core::Interface as _;

use super::traits::{
    AsyncIo2Trait, AsyncIoTrait, BrowseServerAddressSpaceTrait, CommonTrait,
    ConnectionPointContainerTrait, DataObjectTrait, GroupStateMgtTrait, ItemMgtTrait,
    ItemPropertiesTrait, PublicGroupStateMgtTrait, ServerPublicGroupsTrait, ServerTrait,
    SyncIoTrait,
};

pub struct Server {
    pub(crate) server: opc_da_bindings::IOPCServer,
    pub(crate) common: opc_da_bindings::IOPCCommon,
    pub(crate) connection_point_container: windows::Win32::System::Com::IConnectionPointContainer,
    pub(crate) item_properties: opc_da_bindings::IOPCItemProperties,
    pub(crate) server_public_groups: Option<opc_da_bindings::IOPCServerPublicGroups>,
    pub(crate) browse_server_address_space: Option<opc_da_bindings::IOPCBrowseServerAddressSpace>,
}

impl TryFrom<windows_core::IUnknown> for Server {
    type Error = windows_core::Error;

    fn try_from(value: windows_core::IUnknown) -> windows_core::Result<Self> {
        Ok(Self {
            server: value.cast()?,
            common: value.cast()?,
            connection_point_container: value.cast()?,
            item_properties: value.cast()?,
            server_public_groups: value.cast().ok(),
            browse_server_address_space: value.cast().ok(),
        })
    }
}

impl ServerTrait<Group> for Server {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCServer> {
        Ok(&self.server)
    }
}

impl CommonTrait for Server {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCCommon> {
        Ok(&self.common)
    }
}

impl ConnectionPointContainerTrait for Server {
    fn interface(
        &self,
    ) -> windows_core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        Ok(&self.connection_point_container)
    }
}

impl ItemPropertiesTrait for Server {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCItemProperties> {
        Ok(&self.item_properties)
    }
}

impl ServerPublicGroupsTrait for Server {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCServerPublicGroups> {
        self.server_public_groups.as_ref().ok_or_else(|| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCServerPublicGroups not supported",
            )
        })
    }
}

impl BrowseServerAddressSpaceTrait for Server {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCBrowseServerAddressSpace> {
        self.browse_server_address_space.as_ref().ok_or_else(|| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCBrowseServerAddressSpace not supported",
            )
        })
    }
}

pub struct Group {
    pub(crate) item_mgt: opc_da_bindings::IOPCItemMgt,
    pub(crate) group_state_mgt: opc_da_bindings::IOPCGroupStateMgt,
    pub(crate) public_group_state_mgt: Option<opc_da_bindings::IOPCPublicGroupStateMgt>,
    pub(crate) sync_io: opc_da_bindings::IOPCSyncIO,
    pub(crate) async_io: Option<opc_da_bindings::IOPCAsyncIO>,
    pub(crate) async_io2: opc_da_bindings::IOPCAsyncIO2,
    pub(crate) connection_point_container: windows::Win32::System::Com::IConnectionPointContainer,
    pub(crate) data_object: Option<windows::Win32::System::Com::IDataObject>,
}

impl TryFrom<windows_core::IUnknown> for Group {
    type Error = windows_core::Error;

    fn try_from(value: windows_core::IUnknown) -> windows_core::Result<Self> {
        Ok(Self {
            item_mgt: value.cast()?,
            group_state_mgt: value.cast()?,
            public_group_state_mgt: value.cast().ok(),
            sync_io: value.cast()?,
            async_io: value.cast().ok(),
            async_io2: value.cast()?,
            connection_point_container: value.cast()?,
            data_object: value.cast().ok(),
        })
    }
}

impl ItemMgtTrait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCItemMgt> {
        Ok(&self.item_mgt)
    }
}

impl GroupStateMgtTrait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCGroupStateMgt> {
        Ok(&self.group_state_mgt)
    }
}

impl PublicGroupStateMgtTrait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCPublicGroupStateMgt> {
        self.public_group_state_mgt.as_ref().ok_or_else(|| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCPublicGroupStateMgt not supported",
            )
        })
    }
}

impl SyncIoTrait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCSyncIO> {
        Ok(&self.sync_io)
    }
}

impl AsyncIoTrait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCAsyncIO> {
        self.async_io.as_ref().ok_or_else(|| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCAsyncIO not supported",
            )
        })
    }
}

impl AsyncIo2Trait for Group {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCAsyncIO2> {
        Ok(&self.async_io2)
    }
}

impl ConnectionPointContainerTrait for Group {
    fn interface(
        &self,
    ) -> windows_core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        Ok(&self.connection_point_container)
    }
}

impl DataObjectTrait for Group {
    fn interface(&self) -> windows_core::Result<&windows::Win32::System::Com::IDataObject> {
        self.data_object.as_ref().ok_or_else(|| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IDataObject not supported",
            )
        })
    }
}
