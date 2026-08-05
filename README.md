# Rust OPC Classic

Rust bindings and ownership-aware client/server facades for OPC Classic on
Windows. The workspace targets Rust 1.97, edition 2024.

## Workspace

- `opc_classic_utils`: COM apartment, task-memory ownership, transactional
  output builders, panic-safe ABI helpers, class factory and local-server class
  registration.
- `opc_comn_bindings`: generated OPC Common ABI plus safe common/server-list
  clients and server adapters.
- `opc_da_bindings`: generated Data Access ABI plus safe server/group/property
  clients and `IOPCServer`/group adapters.
- `opc_ae_bindings`: generated Alarms & Events ABI plus safe query/subscription
  clients and an `IOPCEventServer` adapter.
- `opc_hda_bindings`: generated Historical Data Access ABI plus safe metadata,
  handle and raw-read clients and `IOPCHDA_Server`/`IOPCHDA_SyncRead` adapters.

Each bindings crate keeps generated code in `src/bindings.rs`. Application code
should normally use its `client` or `server` module; the generated symbols are
still re-exported for unsupported or vendor-specific interfaces.

## Ownership model

Input strings and arrays are Rust-owned borrowed values. COM output memory is
adopted by an explicit owner:

- `OwnedPwstr` owns one task-allocated string.
- `CoTaskMemArray<T, C>` owns a task-allocated array and a cleanup policy for
  nested values.
- `CoTaskMemOut<T>` null-initializes an ABI out pointer before adopting it.
- `CoTaskMemArrayBuilder<T, C>` constructs server outputs transactionally and
  rolls back the initialized prefix after an error or panic.

Owners are not shallow-cloneable. `VARIANT` values use Rust drop glue, string
arrays free each string, and HDA item arrays use method-specific nested cleanup.

## Minimal client setup

```rust,no_run
use opc_classic_utils::ComApartment;
use opc_da_bindings::client::DaClient;

let apartment = ComApartment::mta()?;
let client = DaClient::connect_prog_id(&apartment, "Vendor.OPC.Server")?;
let status = client.status()?;
println!("{}", status.vendor_info);
# Ok::<(), windows_core::Error>(())
```

## Minimal local-server activation

```rust,no_run
use opc_classic_utils::ComApartment;
use opc_classic_utils::server::{ClassFactory, LocalClassRegistration};
use windows_core::{GUID, IUnknown};

# fn make_server() -> windows_core::Result<IUnknown> { unimplemented!() }
let apartment = ComApartment::mta()?;
let class_id = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
let factory = ClassFactory::new(make_server);
let registration = LocalClassRegistration::register(&apartment, &class_id, factory)?;

// Run the application's shutdown/message loop while `registration` is alive.
registration.revoke()?;
# Ok::<(), windows_core::Error>(())
```

The higher-level adapters intentionally return `E_NOTIMPL` for capabilities not
represented by their current Rust service traits. The unsafe generated ABI
remains available when those optional interfaces are needed.

## License notice

This personal project is not affiliated with the OPC Foundation. Files under
the bindings crates' `.metadata` directories originate from the OPC Foundation
and remain subject to its license terms.
