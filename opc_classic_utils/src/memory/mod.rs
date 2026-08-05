//! Ownership-aware helpers for memory crossing an OPC Classic COM boundary.
//!
//! Input arguments are normally borrowed for the duration of a COM call. Output
//! allocations are owned by the receiver and use the deallocator required by
//! their ABI type. The types in this module model that ownership directly.

mod array;
mod out;
mod wide;

pub use array::{
    Cleanup, CoTaskMemArray, CoTaskMemArrayBuilder, DropElements, FreePwstrElements, NoCleanup,
};
pub use out::{CoTaskMemArrayOut, CoTaskMemOut};
pub use wide::{OwnedPwstr, WideCString, WideStringError};
