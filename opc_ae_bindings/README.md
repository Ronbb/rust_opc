# OPC Alarms & Events bindings

`src/bindings.rs` contains private generated code for the unsafe Windows ABI.
Use `client::AeClient` for
status, categories, conditions and subscriptions. `server::AeServer` implements
the query and enable/disable portion of `IOPCEventServer` from the safe
`AeService` contract.

The `client` and `server` APIs expose only standard-library types and types from
`opc_classic_types`, including `ComObject`, `Guid`, `Timestamp`, `Error`, and
`AeServerState`. Windows interface types remain confined to the generated ABI
layer.

The generated bindings were made crate-private in 0.4.0. Code that imported
`opc_ae_bindings::IOPC*` or generated tag types must migrate to the safe
client/server facades or maintain its own ABI layer. `AeServer::new_with_common`
composes an AE service and OPC Common service on one COM identity.
