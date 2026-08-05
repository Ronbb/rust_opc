# OPC Data Access bindings

`src/bindings.rs` contains private generated code for the unsafe Windows ABI. The `client` module provides
`DaClient`, `DaGroup`, item management, synchronous I/O and property APIs. The
`server` module provides Rust service traits and task-memory marshalling
through `DaServer`. Both modules expose only `opc_classic_types`, `core`, and
`std` types; HRESULT, FILETIME, VARIANT, GUID, and COM interfaces are converted
inside the private ABI layer.

The generated bindings were made crate-private in 0.4.0. Code that imported
`opc_da_bindings::IOPC*` or `tagOPC*` symbols must migrate to the safe
client/server facades or maintain its own ABI layer. `DaServer::new_with_common`
composes a DA service and OPC Common service on one COM identity.
