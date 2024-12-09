use windows_core::Interface as _;

use super::traits::{
    AsyncIoTrait, BrowseServerAddressSpaceTrait, DataObjectTrait, GroupStateMgtTrait, ItemMgtTrait,
    PublicGroupStateMgtTrait, ServerPublicGroupsTrait, ServerTrait, SyncIoTrait,
};

pub struct Server {
    pub(crate) server: opc_da_bindings::IOPCServer,
    pub(crate) server_public_groups: Option<opc_da_bindings::IOPCServerPublicGroups>,
    pub(crate) browse_server_address_space: Option<opc_da_bindings::IOPCBrowseServerAddressSpace>,
}

impl TryFrom<windows_core::IUnknown> for Server {
    type Error = windows_core::Error;

    fn try_from(value: windows_core::IUnknown) -> windows_core::Result<Self> {
        Ok(Self {
            server: value.cast()?,
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
    pub(crate) async_io: opc_da_bindings::IOPCAsyncIO,
    pub(crate) data_object: windows::Win32::System::Com::IDataObject,
}

impl TryFrom<windows_core::IUnknown> for Group {
    type Error = windows_core::Error;

    fn try_from(value: windows_core::IUnknown) -> windows_core::Result<Self> {
        Ok(Self {
            item_mgt: value.cast()?,
            group_state_mgt: value.cast()?,
            public_group_state_mgt: value.cast().ok(),
            sync_io: value.cast()?,
            async_io: value.cast()?,
            data_object: value.cast()?,
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
        Ok(&self.async_io)
    }
}

impl DataObjectTrait for Group {
    fn interface(&self) -> windows_core::Result<&windows::Win32::System::Com::IDataObject> {
        Ok(&self.data_object)
    }
}
