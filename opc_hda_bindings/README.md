# OPC Historical Data Access bindings

`src/bindings.rs` is generated unsafe ABI code. Use `client::HdaClient` for
ownership-aware metadata, handles and historical reads.
`server::HdaServerAdapter` implements metadata, handle management and raw reads
with transactional cleanup for every nested HDA item allocation.
