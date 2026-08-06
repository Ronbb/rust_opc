# OPC Common bindings

`src/bindings.rs` is generated unsafe ABI code. Use `client` for ownership-aware
wrappers and `server` for Rust service contracts and COM adapters.

The public `client` and `server` APIs use `opc_classic_types` plus ordinary
`core`/`std` types. They do not expose `windows` or `windows-core` values. Pass
an owning `ComObject` to a client's `from_object` constructor, and obtain the
same Windows-independent handle from a server or client with `object()`.
