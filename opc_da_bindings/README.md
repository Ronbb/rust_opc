# OPC Data Access bindings

`src/bindings.rs` is private generated unsafe ABI code. The `client` module provides
`DaClient`, `DaGroup`, item management, synchronous I/O and property APIs. The
`server` module provides Rust service traits and task-memory marshalling
through `DaServer`. Both modules expose only `opc_classic_types`, `core`, and
`std` types; HRESULT, FILETIME, VARIANT, GUID, and COM interfaces are converted
inside the private ABI layer.
