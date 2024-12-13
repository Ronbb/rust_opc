//! OPC DA 3.0 server and group implementations.
//!
//! This module implements the OPC DA 3.0 interfaces for servers and groups,
//! providing access to the latest features of the OPC DA specification.

use windows::core::Interface as _;

use super::{
    traits::{
        AsyncIo2Trait, AsyncIo3Trait, BrowseTrait, CommonTrait, ConnectionPointContainerTrait,
        GroupStateMgt2Trait, GroupStateMgtTrait, ItemDeadbandMgtTrait, ItemIoTrait, ItemMgtTrait,
        ItemSamplingMgtTrait, ServerTrait, SyncIo2Trait, SyncIoTrait,
    },
    ClientTrait,
};

/// Client for OPC DA 3.0 servers.
pub struct Client;

impl ClientTrait<Server> for Client {
    const CATALOG_ID: windows::core::GUID = opc_da_bindings::CATID_OPCDAServer30::IID;
}

/// An OPC DA 3.0 server implementation.
///
/// Provides access to OPC DA 3.0 server interfaces including:
/// - `IOPCServer` for basic server operations
/// - `IOPCCommon` for server status and locale management
/// - `IOPCBrowse` for browsing the server address space
/// - `IOPCItemIO` for direct item read/write operations
///
/// # Example
/// ```no_run
/// use opc_da::client::v3::Server;
/// // Create server from IUnknown interface
/// let server: Server = unknown.try_into()?;
/// ```
pub struct Server {
    pub(crate) server: opc_da_bindings::IOPCServer,
    pub(crate) common: opc_da_bindings::IOPCCommon,
    pub(crate) connection_point_container: windows::Win32::System::Com::IConnectionPointContainer,
    pub(crate) browse: opc_da_bindings::IOPCBrowse,
    pub(crate) item_io: opc_da_bindings::IOPCItemIO,
}

impl TryFrom<windows::core::IUnknown> for Server {
    type Error = windows::core::Error;

    fn try_from(value: windows::core::IUnknown) -> windows::core::Result<Self> {
        Ok(Self {
            server: value.cast()?,
            common: value.cast()?,
            connection_point_container: value.cast()?,
            browse: value.cast()?,
            item_io: value.cast()?,
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
        Ok(&self.common)
    }
}

impl ConnectionPointContainerTrait for Server {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        Ok(&self.connection_point_container)
    }
}

impl BrowseTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCBrowse> {
        Ok(&self.browse)
    }
}

impl ItemIoTrait for Server {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemIO> {
        Ok(&self.item_io)
    }
}

/// An OPC DA 3.0 group implementation.
///
/// Provides access to OPC DA 3.0 group interfaces including:
/// - `IOPCItemMgt` for item management
/// - `IOPCGroupStateMgt` and `IOPCGroupStateMgt2` for group state management
/// - `IOPCSyncIO` and `IOPCSyncIO2` for synchronous operations
/// - `IOPCAsyncIO2` and `IOPCAsyncIO3` for asynchronous operations
/// - `IOPCItemSamplingMgt` for item sampling control
/// - `IOPCItemDeadbandMgt` for deadband management
#[derive(Debug)]
pub struct Group {
    pub(crate) item_mgt: opc_da_bindings::IOPCItemMgt,
    pub(crate) group_state_mgt: opc_da_bindings::IOPCGroupStateMgt,
    pub(crate) group_state_mgt2: opc_da_bindings::IOPCGroupStateMgt2,
    pub(crate) sync_io: opc_da_bindings::IOPCSyncIO,
    pub(crate) sync_io2: opc_da_bindings::IOPCSyncIO2,
    pub(crate) async_io2: opc_da_bindings::IOPCAsyncIO2,
    pub(crate) async_io3: opc_da_bindings::IOPCAsyncIO3,
    pub(crate) item_sampling_mgt: Option<opc_da_bindings::IOPCItemSamplingMgt>,
    pub(crate) item_deadband_mgt: opc_da_bindings::IOPCItemDeadbandMgt,
    pub(crate) connection_point_container: windows::Win32::System::Com::IConnectionPointContainer,
}

impl TryFrom<windows::core::IUnknown> for Group {
    type Error = windows::core::Error;

    fn try_from(value: windows::core::IUnknown) -> windows::core::Result<Self> {
        Ok(Self {
            item_mgt: value.cast()?,
            group_state_mgt: value.cast()?,
            group_state_mgt2: value.cast()?,
            sync_io: value.cast()?,
            sync_io2: value.cast()?,
            async_io2: value.cast()?,
            async_io3: value.cast()?,
            item_deadband_mgt: value.cast()?,
            item_sampling_mgt: value.cast().ok(),
            connection_point_container: value.cast()?,
        })
    }
}

impl ItemMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemMgt> {
        Ok(&self.item_mgt)
    }
}

impl GroupStateMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt> {
        Ok(&self.group_state_mgt)
    }
}

impl GroupStateMgt2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt2> {
        Ok(&self.group_state_mgt2)
    }
}

impl SyncIoTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO> {
        Ok(&self.sync_io)
    }
}

impl SyncIo2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO2> {
        Ok(&self.sync_io2)
    }
}

impl AsyncIo2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO2> {
        Ok(&self.async_io2)
    }
}

impl AsyncIo3Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO3> {
        Ok(&self.async_io3)
    }
}

impl ItemDeadbandMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemDeadbandMgt> {
        Ok(&self.item_deadband_mgt)
    }
}

impl ItemSamplingMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemSamplingMgt> {
        self.item_sampling_mgt.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCItemSamplingMgt not supported",
            )
        })
    }
}

impl ConnectionPointContainerTrait for Group {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        Ok(&self.connection_point_container)
    }
}
