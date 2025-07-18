//! Memory Management Comparison Example
//!
//! This example demonstrates the key differences between `CalleeAllocatedArray` and `CalleeAllocatedPtrArray`
//! to help users understand when to use each type.

use opc_classic_utils::memory::{
    CalleeAllocatedArray, CalleeAllocatedPtrArray, CallerAllocatedArray, CallerAllocatedPtrArray,
};
use std::ptr;

fn main() {
    println!("=== OPC Classic Memory Management Comparison ===\n");

    // 1. Value array comparison
    println!("1. Value Array Comparison:");
    compare_value_arrays();
    println!();

    // 2. Pointer array comparison
    println!("2. Pointer Array Comparison:");
    compare_pointer_arrays();
    println!();

    // 3. Real-world usage scenarios
    println!("3. Real-World Usage Scenarios:");
    demonstrate_real_world_usage();
}

fn compare_value_arrays() {
    println!("  CalleeAllocatedArray (Callee-allocated):");
    println!("  - Manages the array container itself");
    println!("  - Does not free array elements");
    println!("  - Use case: Simple data type arrays");

    // Simulate COM-allocated value array - use null pointer to avoid actual allocation
    let array = CalleeAllocatedArray::from_raw(ptr::null_mut::<i32>(), 0);
    println!("  - Created empty array (length: {})", array.len());
    // Automatically frees array container

    println!("\n  CallerAllocatedArray (Caller-allocated):");
    println!("  - Caller responsible for allocation and deallocation");
    println!("  - Full control over memory lifecycle");
    println!("  - Use case: Scenarios requiring precise memory control");

    let caller_array: CallerAllocatedArray<i32> = CallerAllocatedArray::allocate(5).unwrap();
    println!("  - Allocated array (length: {})", caller_array.len());
    // Caller responsible for deallocation
}

fn compare_pointer_arrays() {
    println!("  CalleeAllocatedPtrArray (Callee-allocated):");
    println!("  - Manages array container and each pointer element");
    println!("  - Frees array container + each pointer's memory");
    println!("  - Use case: COM interface pointer arrays");

    // Simulate COM-allocated pointer array - use null pointer to avoid actual allocation
    let ptr_array = CalleeAllocatedPtrArray::from_raw(ptr::null_mut::<*mut i32>(), 0);
    println!(
        "  - Created empty pointer array (length: {})",
        ptr_array.len()
    );
    // Automatically frees array container and each pointer

    println!("\n  CallerAllocatedPtrArray (Caller-allocated):");
    println!("  - Caller responsible for allocation and deallocation");
    println!("  - Full control over pointer array lifecycle");
    println!("  - Use case: Scenarios requiring precise pointer memory control");

    let caller_ptr_array: CallerAllocatedPtrArray<i32> =
        CallerAllocatedPtrArray::allocate(3).unwrap();
    println!(
        "  - Allocated pointer array (length: {})",
        caller_ptr_array.len()
    );
    // Caller responsible for deallocation
}

fn demonstrate_real_world_usage() {
    println!("  Scenario 1: Batch reading OPC item values");
    println!("  - Use CalleeAllocatedArray to receive value arrays");
    println!("  - COM server allocates memory, client frees container");

    let value_array = CalleeAllocatedArray::from_raw(ptr::null_mut::<f64>(), 0);
    println!("  - Read {} values", value_array.len());

    println!("\n  Scenario 2: Getting OPC server interface list");
    println!("  - Use CalleeAllocatedPtrArray to receive interface pointer arrays");
    println!("  - COM server allocates memory and interface pointers, client frees all resources");

    let interface_array = CalleeAllocatedPtrArray::from_raw(ptr::null_mut::<*mut ()>(), 0);
    println!("  - Retrieved {} interface pointers", interface_array.len());

    println!("\n  Scenario 3: Custom memory management");
    println!("  - Use CallerAllocatedArray for precise memory control");
    println!("  - Suitable for performance-critical scenarios");

    let custom_array: CallerAllocatedArray<f64> = CallerAllocatedArray::allocate(4).unwrap();
    println!("  - Custom array (length: {})", custom_array.len());

    println!("\n  Best Practices Summary:");
    println!("  - Simple data: Use CalleeAllocatedArray");
    println!("  - Interface pointers: Use CalleeAllocatedPtrArray");
    println!("  - Performance optimization: Use CallerAllocatedArray/CallerAllocatedPtrArray");
    println!("  - Avoid manual memory management errors");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_memory_management_differences() {
        // This test demonstrates the different memory management behaviors

        // Test CalleeAllocatedArray - should only free the container
        {
            let _array = CalleeAllocatedArray::<i32>::default();
            // When _array goes out of scope, only the container is freed
        }

        // Test CalleeAllocatedPtrArray - should free both container and elements
        {
            let _ptr_array = CalleeAllocatedPtrArray::<i32>::default();
            // When _ptr_array goes out of scope, both container and elements are freed
        }

        // Both should complete without memory leaks
    }
}
