# OPC Historical Data Access bindings

`src/bindings.rs` is private generated unsafe ABI code. Use `client::HdaClient` for
ownership-aware metadata, handles and historical reads.
`server::HdaServer` implements metadata, handle management and raw reads
with transactional cleanup for every nested HDA item allocation.

The public client and server contracts use `Timestamp`, `Value`, `ErrorCode`,
`ComObject`, and other types from `opc_classic_types`; Windows ABI types remain
private conversion details.
