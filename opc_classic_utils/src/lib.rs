//! Shared safe primitives for OPC Classic COM clients and servers.

pub mod com;
pub mod memory;
pub mod server;

pub use com::{ApartmentModel, ComApartment};
pub use memory::*;
pub use opc_classic_types::{
    ClassContext, ComObject, Error, ErrorCode, ErrorKind, Guid, Result, Timestamp, Value, ValueType,
};
