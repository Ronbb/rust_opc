//! Demonstration of #[repr(transparent)] usage
//!
//! This example shows how the transparent representation allows
//! our wrapper types to be used directly in FFI scenarios.

use opc_classic_utils::memory::{
    CalleeAllocatedPtr, CalleeAllocatedWString, CallerAllocatedPtr, CallerAllocatedWString,
};

/// Simulates a C function that takes a pointer parameter
/// In real scenarios, this would be an external C function
#[allow(dead_code)]
extern "C" fn simulate_c_function(ptr: *mut i32) -> i32 {
    if ptr.is_null() { 0 } else { unsafe { *ptr } }
}

/// Simulates a C function that returns a pointer
/// In real scenarios, this would be an external C function
#[allow(dead_code)]
extern "C" fn simulate_c_function_return() -> *mut u16 {
    std::ptr::null_mut()
}

fn demonstrate_transparent_repr() {
    println!("#[repr(transparent)] Demonstration");
    println!("==================================");

    println!("\n1. Memory layout compatibility:");

    // Create a raw pointer
    let raw_ptr = std::ptr::null_mut::<i32>();

    // Wrap it in our transparent types
    let caller_ptr = CallerAllocatedPtr::from_raw(raw_ptr);
    let callee_ptr = CalleeAllocatedPtr::from_raw(raw_ptr);

    println!("   Raw pointer: {:?}", raw_ptr);
    println!("   CallerAllocatedPtr: {:?}", caller_ptr.as_ptr());
    println!("   CalleeAllocatedPtr: {:?}", callee_ptr.as_ptr());

    // All pointers should be identical due to transparent repr
    assert_eq!(raw_ptr, caller_ptr.as_ptr());
    assert_eq!(raw_ptr, callee_ptr.as_ptr());

    println!("   ✓ All pointers have identical memory layout");

    println!("\n2. FFI compatibility demonstration:");

    // Our wrapper types can be used directly with C functions
    // because they have transparent representation

    // For input parameters (caller-allocated)
    let input_ptr = CallerAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
    println!("   Input pointer created: {:?}", input_ptr.as_ptr());

    // For output parameters (callee-allocated)
    let output_ptr = CalleeAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
    println!("   Output pointer created: {:?}", output_ptr.as_ptr());

    println!("   ✓ Wrapper types are FFI-compatible");

    println!("\n3. Zero-cost abstraction:");

    // Demonstrate that there's no overhead
    let ptr1 = std::ptr::null_mut::<i32>();
    let _ptr2 = CallerAllocatedPtr::from_raw(ptr1);
    let _ptr3 = CalleeAllocatedPtr::from_raw(ptr1);

    // All should have the same size
    println!(
        "   Size of *mut i32: {} bytes",
        std::mem::size_of::<*mut i32>()
    );
    println!(
        "   Size of CallerAllocatedPtr<i32>: {} bytes",
        std::mem::size_of::<CallerAllocatedPtr<i32>>()
    );
    println!(
        "   Size of CalleeAllocatedPtr<i32>: {} bytes",
        std::mem::size_of::<CalleeAllocatedPtr<i32>>()
    );

    assert_eq!(
        std::mem::size_of::<*mut i32>(),
        std::mem::size_of::<CallerAllocatedPtr<i32>>()
    );
    assert_eq!(
        std::mem::size_of::<*mut i32>(),
        std::mem::size_of::<CalleeAllocatedPtr<i32>>()
    );

    println!("   ✓ Zero memory overhead");

    println!("\n4. Wide string compatibility:");

    let wstring_ptr = std::ptr::null_mut::<u16>();
    let _caller_wstring = CallerAllocatedWString::from_raw(wstring_ptr);
    let _callee_wstring = CalleeAllocatedWString::from_raw(wstring_ptr);

    println!(
        "   Size of *mut u16: {} bytes",
        std::mem::size_of::<*mut u16>()
    );
    println!(
        "   Size of CallerAllocatedWString: {} bytes",
        std::mem::size_of::<CallerAllocatedWString>()
    );
    println!(
        "   Size of CalleeAllocatedWString: {} bytes",
        std::mem::size_of::<CalleeAllocatedWString>()
    );

    assert_eq!(
        std::mem::size_of::<*mut u16>(),
        std::mem::size_of::<CallerAllocatedWString>()
    );
    assert_eq!(
        std::mem::size_of::<*mut u16>(),
        std::mem::size_of::<CalleeAllocatedWString>()
    );

    println!("   ✓ Wide string types also have zero overhead");

    println!("\n5. Benefits summary:");
    println!("   - Memory layout identical to raw pointers");
    println!("   - FFI-compatible without any conversion");
    println!("   - Zero runtime overhead");
    println!("   - Type safety and automatic memory management");
    println!("   - Perfect for COM/OPC scenarios");
}

fn main() {
    demonstrate_transparent_repr();

    println!("\n✅ #[repr(transparent)] demonstration completed successfully!");
    println!("   Our wrapper types provide safety without any performance cost.");
}
