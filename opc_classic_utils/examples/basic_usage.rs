//! Basic usage example for opc_classic_utils
//!
//! This example demonstrates how to use the memory management utilities
//! provided by opc_classic_utils, showing both caller-allocated and callee-allocated patterns.

use opc_classic_utils::memory::{
    CalleeAllocatedPtr, CalleeAllocatedWString, CallerAllocatedPtr, CallerAllocatedWString,
};

fn main() {
    println!("OPC Classic Utils - Memory Management Patterns");
    println!("==============================================");

    // Example 1: Caller-allocated pointers (input parameters)
    println!("\n1. Caller-allocated pointers (input parameters):");
    println!("   - Caller allocates memory");
    println!("   - Callee (COM function) is responsible for freeing");
    println!("   - Our wrapper does NOT free the memory");

    let caller_ptr: CallerAllocatedPtr<i32> = CallerAllocatedPtr::default();
    println!(
        "   Created caller-allocated pointer, null: {}",
        caller_ptr.is_null()
    );

    // Example 2: Callee-allocated pointers (output parameters)
    println!("\n2. Callee-allocated pointers (output parameters):");
    println!("   - Callee (COM function) allocates memory");
    println!("   - Caller is responsible for freeing using CoTaskMemFree");
    println!("   - Our wrapper automatically frees the memory when dropped");

    let callee_ptr: CalleeAllocatedPtr<i32> = CalleeAllocatedPtr::default();
    println!(
        "   Created callee-allocated pointer, null: {}",
        callee_ptr.is_null()
    );

    // Example 3: Caller-allocated wide strings
    println!("\n3. Caller-allocated wide strings:");
    let caller_wstring = CallerAllocatedWString::default();
    println!(
        "   Created caller-allocated wide string, null: {}",
        caller_wstring.is_null()
    );

    // Example 4: Callee-allocated wide strings
    println!("\n4. Callee-allocated wide strings:");
    let callee_wstring = CalleeAllocatedWString::default();
    println!(
        "   Created callee-allocated wide string, null: {}",
        callee_wstring.is_null()
    );

    // Example 5: Demonstrating automatic cleanup for callee-allocated memory
    println!("\n5. Demonstrating automatic cleanup for callee-allocated memory:");
    {
        // This would normally be a pointer returned from a COM function
        let _ptr = CalleeAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
        println!(
            "   Created callee-allocated pointer, will be automatically freed when scope ends"
        );
    } // _ptr is automatically dropped and memory freed here

    // Example 6: Demonstrating no cleanup for caller-allocated memory
    println!("\n6. Demonstrating no cleanup for caller-allocated memory:");
    {
        // This would normally be a pointer allocated by the caller
        let _ptr = CallerAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
        println!("   Created caller-allocated pointer, callee will be responsible for freeing");
    } // _ptr is dropped but memory is NOT freed (callee's responsibility)

    println!("\n7. Example completed successfully!");
    println!("   Memory management follows COM conventions:");
    println!("   - Caller-allocated: callee frees");
    println!("   - Callee-allocated: caller frees (handled automatically)");
}
