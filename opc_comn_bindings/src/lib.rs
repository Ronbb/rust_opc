mod abi;
mod bindings;
pub mod client;
pub mod server;

// Keep generated COM ABI symbols inside this crate; use the safe client/server
// facades for application code.
pub(crate) use bindings::*;
