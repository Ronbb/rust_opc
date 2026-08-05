# OPC Data Access bindings

`src/bindings.rs` is generated unsafe ABI code. The `client` module provides
`DaClient`, `DaGroup`, item management, synchronous I/O and property APIs. The
`server` module provides Rust service traits and task-memory marshalling
adapters.
