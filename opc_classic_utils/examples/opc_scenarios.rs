//! OPC scenarios example for opc_classic_utils
//!
//! This example demonstrates how to use the memory management utilities
//! in typical OPC scenarios, showing the correct patterns for different situations.

use opc_classic_utils::memory::{
    CalleeAllocatedPtr, CalleeAllocatedWString, CallerAllocatedPtr, CallerAllocatedWString,
};

/// Simulates an OPC server method that takes input parameters (caller-allocated)
/// and returns output parameters (callee-allocated)
fn simulate_opc_server_method(
    input_string: &CallerAllocatedWString, // Caller allocates, server frees
    input_data: &CallerAllocatedPtr<i32>,  // Caller allocates, server frees
) -> (
    CalleeAllocatedWString,  // Server allocates, caller frees
    CalleeAllocatedPtr<i32>, // Server allocates, caller frees
) {
    println!("  Server: Processing input parameters...");
    println!("  Server: Input string is null: {}", input_string.is_null());
    println!("  Server: Input data is null: {}", input_data.is_null());

    // Simulate server allocating memory for output
    println!("  Server: Allocating memory for output parameters...");

    // In a real scenario, the server would allocate memory using COM functions
    // and return pointers to that memory
    (
        CalleeAllocatedWString::default(), // Simulated server-allocated string
        CalleeAllocatedPtr::default(),     // Simulated server-allocated data
    )
}

/// Simulates an OPC client calling a server method
fn simulate_opc_client_call() {
    println!("OPC Client-Server Memory Management Example");
    println!("===========================================");

    println!("\n1. Client preparing input parameters (caller-allocated):");

    // Client allocates memory for input parameters
    let client_input_string = CallerAllocatedWString::from_raw(std::ptr::null_mut());
    let client_input_data = CallerAllocatedPtr::from_raw(std::ptr::null_mut());

    println!("   Client: Allocated memory for input string and data");
    println!("   Client: These will be freed by the server");

    println!("\n2. Client calling server method:");

    // Call the server method
    let (server_output_string, server_output_data) =
        simulate_opc_server_method(&client_input_string, &client_input_data);

    println!("\n3. Client processing output parameters (callee-allocated):");
    println!(
        "   Client: Received output string, null: {}",
        server_output_string.is_null()
    );
    println!(
        "   Client: Received output data, null: {}",
        server_output_data.is_null()
    );
    println!("   Client: These will be automatically freed when they go out of scope");

    // The server_output_string and server_output_data will be automatically
    // freed when they go out of scope at the end of this function
}

/// Demonstrates the difference between input and output parameters
fn demonstrate_parameter_differences() {
    println!("\n\nOPC Parameter Type Differences");
    println!("===============================");

    println!("\nInput Parameters (Caller-allocated):");
    println!("  - Client allocates memory");
    println!("  - Server uses the memory");
    println!("  - Server frees the memory");
    println!("  - Our wrapper does NOT free (server's responsibility)");

    println!("\nOutput Parameters (Callee-allocated):");
    println!("  - Server allocates memory");
    println!("  - Client receives the memory");
    println!("  - Client frees the memory");
    println!("  - Our wrapper automatically frees (client's responsibility)");

    println!("\nKey Benefits:");
    println!("  - Clear ownership semantics");
    println!("  - Prevents double-free errors");
    println!("  - Prevents memory leaks");
    println!("  - Follows COM conventions exactly");
}

fn main() {
    simulate_opc_client_call();
    demonstrate_parameter_differences();

    println!("\n\nExample completed successfully!");
    println!("Memory management follows OPC/COM conventions correctly.");
}
