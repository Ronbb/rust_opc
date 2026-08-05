# OPC Alarms & Events bindings

`src/bindings.rs` is generated unsafe ABI code. Use `client::AeClient` for
status, categories, conditions and subscriptions. `server::AeServerAdapter`
implements the query and enable/disable portion of `IOPCEventServer` from the
safe `AeService` contract.
