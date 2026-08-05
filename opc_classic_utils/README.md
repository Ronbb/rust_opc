# OPC Classic utilities

This crate contains the ownership and COM-apartment primitives shared by the
OPC Classic client and server facades.

## Ownership model

COM input parameters are borrowed for the duration of a call. Use
`WideCString` for strings and ordinary Rust slices for arrays. Do not allocate
an input buffer solely because an IDL parameter is a pointer.

COM output parameters are adopted by an owning guard:

```rust
use opc_classic_utils::{CoTaskMemOut, NoCleanup, OwnedPwstr};

let mut output = CoTaskMemOut::<u32>::new();
// unsafe { raw_com_method(output.as_mut_ptr()) };
let values = unsafe { output.into_array(3, NoCleanup) }?;
println!("{} values", values.len());

let mut text = CoTaskMemOut::<u16>::new();
// unsafe { raw_string_method(text.as_mut_ptr()) };
let text: OwnedPwstr = unsafe { text.into_pwstr() };
println!("{}", text.to_string_lossy());
# Ok::<(), windows_core::Error>(())
```

`CoTaskMemArray<T, C>` always releases its outer allocation with
`CoTaskMemFree`. The cleanup policy `C` decides what happens to initialized
elements. Use `DropElements` for `VARIANT` values, `FreePwstrElements` for
arrays of task-allocated strings, and `NoCleanup` only for plain ABI values or
when a method-specific decoder owns nested cleanup.

Owning guards are deliberately not `Clone`, and adopting a raw pointer is an
`unsafe` operation. This prevents accidental aliases and allocator mismatches.

## Server activation

`server::ClassFactory` turns a Rust closure returning `IUnknown` into an
`IClassFactory`. `server::LocalClassRegistration` registers it with COM and
revokes the registration on drop. Its borrow of `ComApartment` prevents the
apartment from being uninitialized before the registration is revoked.
