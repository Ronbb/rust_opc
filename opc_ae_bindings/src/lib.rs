mod abi;
mod bindings;
pub mod client;
pub mod server;
mod types;

// The generated Windows ABI is an implementation detail. Public client/server
// facades expose only `opc_classic_types` and standard-library types.
pub(crate) use bindings::*;
pub use types::{AeServerState, BrowseDirection};
