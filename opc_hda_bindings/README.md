# OPC Historical Data Access bindings

`src/bindings.rs` contains private generated code for the unsafe Windows ABI. It
is not a public API. Use `client::HdaClient` for ownership-aware metadata,
handles and historical reads.
`server::HdaServer` implements metadata, handle management and raw reads
with transactional cleanup for every nested HDA item allocation. Use
`HdaServer::new_with_common` when the same COM identity should expose an OPC
Common service; `HdaServer::new` exposes the interface with an explicit
not-configured result for its methods.

The public client and server contracts use `Timestamp`, `Value`, `ErrorCode`,
`ComObject`, and other types from `opc_classic_types`; Windows ABI types remain
private conversion details. The generated bindings were made crate-private in
0.4.0, so code that imported `opc_hda_bindings::IOPCHDA_*` must migrate to the
safe client/server facades (or its own ABI layer).
