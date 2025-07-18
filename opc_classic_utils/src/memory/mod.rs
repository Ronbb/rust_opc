//! Memory management utilities for OPC Classic
//!
//! This module provides automatic memory management for COM objects
//! using `CoTaskMemFree` for cleanup.
//!
//! COM memory management follows two patterns:
//! 1. Caller allocates, callee frees (e.g., input parameters)
//! 2. Callee allocates, caller frees (e.g., output parameters)

pub mod array;
pub mod ptr;
pub mod ptr_array;
pub mod wstring;

// Re-export all public types for convenience
pub use array::{CalleeAllocatedArray, CallerAllocatedArray};
pub use ptr::{CalleeAllocatedPtr, CallerAllocatedPtr};
pub use ptr_array::{CalleeAllocatedPtrArray, CallerAllocatedPtrArray};
pub use wstring::{CalleeAllocatedWString, CallerAllocatedWString};

#[cfg(test)]
mod tests;
