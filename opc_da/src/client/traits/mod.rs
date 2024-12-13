mod async_io;
mod async_io2;
mod async_io3;
mod browse;
mod browse_server_address_space;
/// OPC DA client trait definitions.
///
/// This module contains all trait definitions for interacting with OPC DA servers.
/// The traits are organized by functionality and OPC DA version compatibility:
///
/// Version independent traits:
/// - CommonTrait: Basic server configuration and error handling
/// - ConnectionPointContainerTrait: Event connection management
/// - DataObjectTrait: COM data transfer functionality
///
/// OPC DA 1.0 traits:
/// - AsyncIoTrait: Basic asynchronous operations
/// - SyncIoTrait: Basic synchronous operations
///
/// OPC DA 2.0 traits:
/// - AsyncIo2Trait: Enhanced asynchronous operations
/// - SyncIo2Trait: Enhanced synchronous operations
/// - BrowseServerAddressSpaceTrait: Address space navigation
///
/// OPC DA 3.0 traits:
/// - AsyncIo3Trait: Advanced asynchronous operations
/// - BrowseTrait: Enhanced browsing capabilities
/// - ItemDeadbandMgtTrait: Item deadband management
/// - ItemIoTrait: Direct item access
/// - ItemSamplingMgtTrait: Sampling rate control
/// - GroupStateMgt2Trait: Extended group management
mod client;
mod common;
mod connection_point_container;
mod data_object;
mod group_state_mgt;
mod group_state_mgt2;
mod item_deadband_mgt;
mod item_io;
mod item_mgt;
mod item_properties;
mod item_sampling_mgt;
mod public_group_state_mgt;
mod server;
mod server_public_groups;
mod sync_io;
mod sync_io2;

pub use async_io::*;
pub use async_io2::*;
pub use async_io3::*;
pub use browse::*;
pub use browse_server_address_space::*;
pub use client::*;
pub use common::*;
pub use connection_point_container::*;
pub use data_object::*;
pub use group_state_mgt::*;
pub use group_state_mgt2::*;
pub use item_deadband_mgt::*;
pub use item_io::*;
pub use item_mgt::*;
pub use item_properties::*;
pub use item_sampling_mgt::*;
pub use public_group_state_mgt::*;
pub use server::*;
pub use server_public_groups::*;
pub use sync_io::*;
pub use sync_io2::*;
