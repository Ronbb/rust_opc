# OPC Classic Utils

OPC Classic utilities and common functionality for the rust_opc project.

## Features

- **Dual Memory Management Patterns**: Supports both COM memory management conventions
  - **Caller-allocated**: Caller allocates, callee frees (for input parameters)
  - **Callee-allocated**: Callee allocates, caller frees (for output parameters)
- Automatic memory management with `CoTaskMemFree`
- Common utility structures and traits
- Shared functionality for OPC Classic implementations

## Memory Management Patterns

### Caller-allocated Memory (Input Parameters)

Use these types when **you** allocate memory that will be freed by the COM function:

```rust
use opc_classic_utils::memory::{CallerAllocatedPtr, CallerAllocatedWString};

// For input parameters - you allocate, COM function frees
let input_string = CallerAllocatedWString::from_raw(your_allocated_string);
let input_data = CallerAllocatedPtr::from_raw(your_allocated_data);

// Pass to COM function - it will free the memory
some_com_function(&input_string, &input_data);
// Memory is NOT freed by our wrapper (COM function's responsibility)
```

### Callee-allocated Memory (Output Parameters)

Use these types when the **COM function** allocates memory that you must free:

```rust
use opc_classic_utils::memory::{CalleeAllocatedPtr, CalleeAllocatedWString};

// For output parameters - COM function allocates, you free
let (output_string, output_data) = some_com_function();

// Use the returned data
println!("String: {:?}", output_string.as_ptr());
println!("Data: {:?}", output_data.as_ptr());

// Memory is automatically freed when variables go out of scope
```

## Usage Examples

### Basic Usage

```rust
use opc_classic_utils::memory::{
    CallerAllocatedPtr, CalleeAllocatedPtr,
    CallerAllocatedWString, CalleeAllocatedWString,
};

// Caller-allocated (input parameters)
let input_ptr = CallerAllocatedPtr::from_raw(some_pointer);
// Memory will NOT be freed by our wrapper

// Callee-allocated (output parameters)
let output_ptr = CalleeAllocatedPtr::from_raw(com_returned_pointer);
// Memory will be automatically freed when dropped
```

### OPC Client-Server Scenario

```rust
// Client side - preparing input parameters
let client_input = CallerAllocatedWString::from_raw(client_allocated_string);

// Call server method
let server_output = server_method(&client_input);

// Server output is automatically freed when it goes out of scope
// Client input is NOT freed (server's responsibility)
```

## Type Overview

| Type | Purpose | Memory Management |
|------|---------|-------------------|
| `CallerAllocatedPtr<T>` | Input parameters | Caller allocates, callee frees |
| `CalleeAllocatedPtr<T>` | Output parameters | Callee allocates, caller frees |
| `CallerAllocatedWString` | Input strings | Caller allocates, callee frees |
| `CalleeAllocatedWString` | Output strings | Callee allocates, caller frees |

## Benefits

- **Prevents Memory Leaks**: Automatic cleanup for callee-allocated memory
- **Prevents Double-free**: Clear ownership semantics prevent errors
- **COM Convention Compliance**: Follows standard COM memory management rules
- **Type Safety**: Compile-time guarantees about memory ownership
- **RAII**: Resource management through Rust's ownership system

## License

MIT 