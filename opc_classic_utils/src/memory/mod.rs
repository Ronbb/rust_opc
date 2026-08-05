//! Ownership-aware helpers for memory crossing an OPC Classic COM boundary.
//!
//! Input arguments are normally borrowed for the duration of a COM call. Output
//! allocations are owned by the receiver and use the deallocator required by
//! their ABI type. The types in this module model that ownership directly.
//!
//! A returned array with a fixed, contract-defined length should use
//! [`CoTaskMemArrayOut::new`]. If the server reports the length through a
//! separate out parameter, use [`CoTaskMemArrayOut::new_reported`] and commit
//! the length only after the foreign status has been checked. This keeps a
//! failing server from controlling how many nested pointers are walked during
//! cleanup.

mod array;
mod object;
mod out;
mod wide;

pub use array::{
    Cleanup, CoTaskMemArray, CoTaskMemArrayBuilder, DropElements, FreePwstrElements, NoCleanup,
};
pub use object::CoTaskMemObject;
pub use out::{CoTaskMemArrayOut, CoTaskMemObjectOut, CoTaskMemOut};
pub use wide::{OwnedPwstr, WideCString, WideStringError};
