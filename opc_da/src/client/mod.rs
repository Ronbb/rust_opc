//! OPC DA client implementation.
//!
//! This module provides implementations for OPC DA client functionality across
//! different versions of the specification (1.0, 2.0, and 3.0). It includes:
//!
//! - Version-specific implementations in `v1`, `v2`, and `v3` modules
//! - A unified client interface in the `unified` module
//! - Common traits and memory management utilities

mod iterator;
mod traits;

pub mod unified;
pub mod v1;
pub mod v2;
pub mod v3;

pub use iterator::*;
pub use traits::*;

#[cfg(test)]
mod tests;
