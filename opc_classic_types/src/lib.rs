//! Pure Rust public types shared by OPC Classic clients and servers.
//!
//! This crate deliberately has no dependency on the Windows bindings. COM ABI
//! conversion belongs in the private implementation layer of each bindings
//! crate.

mod com;
mod error;
mod value;

pub use com::{ClassContext, ComObject, Guid, Timestamp};
pub use error::{Error, ErrorCode, ErrorKind, Result};
pub use value::{Value, ValueType};
