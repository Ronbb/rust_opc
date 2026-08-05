# OPC Alarms & Events bindings

`src/bindings.rs` is generated unsafe ABI code. Use `client::AeClient` for
status, categories, conditions and subscriptions. `server::AeServer` implements
the query and enable/disable portion of `IOPCEventServer` from the safe
`AeService` contract.

The `client` and `server` APIs expose only standard-library types and types from
`opc_classic_types`, including `ComObject`, `Guid`, `Timestamp`, `Error`, and
`AeServerState`. Windows interface types remain confined to the generated ABI
layer.
