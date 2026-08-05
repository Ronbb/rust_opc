# Rust OPC Classic

Rust bindings and ownership-aware client/server facades for OPC Classic on
Windows. The workspace targets Rust 1.97, edition 2024.

## Workspace

- `opc_classic_types`: dependency-free public `Guid`, `Timestamp`, `Value`,
  `ComObject`, and OPC error types.
- `opc_classic_abi`: unpublished shared ABI definitions used only to compose
  multiple OPC interfaces on one private COM identity.
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

Each bindings crate keeps generated code in a private `src/bindings.rs` module.
Application code uses its `client` or `server` module. Their public signatures
contain only `opc_classic_types`, `core`, and `std` types; Windows ABI values are
converted inside private boundary modules.

This private visibility is a 0.4.0 breaking change: applications that imported
generated `IOPC*`, `tagOPC*`, or `OPC*` symbols must use the safe facades or
maintain a separate ABI layer. DA, AE, and HDA server adapters expose their
domain interfaces and `IOPCCommon` from the same COM identity; the
`new_with_common` constructors attach application-defined common behavior.

## Ownership model

Input strings and arrays are Rust-owned borrowed values. COM output memory is
adopted by an explicit owner:

- `OwnedPwstr` owns one task-allocated string.
- `CoTaskMemArray<T, C>` owns a task-allocated array and a cleanup policy for
  nested values.
- `CoTaskMemArrayOut<T, C>` distinguishes fixed, trusted lengths from lengths
  reported by a foreign call. Reported lengths are committed only after a
  successful HRESULT; failure cleanup never walks an untrusted count.
- `CoTaskMemObjectOut<T, C>` owns one task-allocated object returned through
  `T**`, including its nested cleanup policy.
- `CoTaskMemOut<T>` null-initializes a single-allocation out pointer.
- `CoTaskMemArrayBuilder<T, C>` constructs server outputs transactionally and
  rolls back the initialized prefix after an error or panic.

Owners are not shallow-cloneable. `VARIANT` values use Rust drop glue, string
arrays free each string, and HDA item arrays use method-specific nested cleanup.

## Minimal client setup

```rust,no_run
use opc_classic_utils::ComApartment;
use opc_da_bindings::client::DaClient;

fn example() -> opc_classic_utils::Result<()> {
    let apartment = ComApartment::mta()?;
    let client = DaClient::connect_prog_id(&apartment, "Vendor.OPC.Server")?;
    let status = client.status()?;
    println!("{}", status.vendor_info);
    Ok(())
}
```

## Minimal local-server activation

```rust,no_run
use opc_classic_utils::{ComApartment, ComObject, Guid, Result};
use opc_classic_utils::server::{ClassFactory, LocalClassRegistration};

fn make_server() -> Result<ComObject> {
    unimplemented!()
}

fn example() -> Result<()> {
    let apartment = ComApartment::mta()?;
    let class_id = Guid::new(
        0x1234_5678,
        0x1234,
        0x1234,
        [0x12, 0x34, 0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc],
    );
    let factory = ClassFactory::new(make_server);
    let registration = LocalClassRegistration::register(&apartment, &class_id, factory)?;

    // Run the application's shutdown/message loop while `registration` is alive.
    registration.revoke()?;
    Ok(())
}
```

The higher-level adapters intentionally return the COM equivalent of
`ErrorCode::NOT_IMPLEMENTED` for capabilities not represented by their current
Rust service traits. Extend the private ABI boundary when adding another
interface instead of leaking Windows types into the public facade.

## License notice

This personal project is not affiliated with the OPC Foundation. Files under
the bindings crates' `.metadata` directories originate from the OPC Foundation
and remain subject to its license terms.
