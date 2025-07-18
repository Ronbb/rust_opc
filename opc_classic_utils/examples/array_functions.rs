//! Array functions demonstration for opc_classic_utils
//!
//! This example shows how to use the array and pointer array functions for
//! COM memory management in OPC Classic scenarios.

use opc_classic_utils::memory::{
    CallerAllocatedArray, CallerAllocatedPtr, CallerAllocatedPtrArray, CallerAllocatedWString,
};
use std::str::FromStr;

#[derive(Debug, Clone, Copy, PartialEq)]
struct OPCItem {
    id: i32,
    value: f64,
    quality: u16,
}

fn demonstrate_array_functions() {
    println!("Array Functions Demonstration");
    println!("============================");

    println!("\n1. Creating arrays from Rust slices:");

    // Create arrays from Rust data
    let item_data = vec![
        OPCItem {
            id: 1,
            value: 10.5,
            quality: 192,
        },
        OPCItem {
            id: 2,
            value: 20.7,
            quality: 192,
        },
        OPCItem {
            id: 3,
            value: 30.2,
            quality: 192,
        },
    ];

    let mut caller_array = CallerAllocatedArray::from_slice(&item_data).unwrap();
    println!(
        "   Created caller-allocated array with {} items",
        caller_array.len()
    );

    // Verify the data was copied correctly
    unsafe {
        let slice = caller_array.as_slice().unwrap();
        assert_eq!(slice, item_data.as_slice());
        println!("   ✓ Data was correctly copied to COM memory");
    }

    println!("\n2. Allocating uninitialized arrays:");

    // Allocate memory for later initialization
    let mut uninit_array = CallerAllocatedArray::<i32>::allocate(5).unwrap();
    println!("   Allocated uninitialized array for 5 i32 values");

    // Initialize the array
    unsafe {
        let mut_slice = uninit_array.as_mut_slice().unwrap();
        for (i, value) in mut_slice.iter_mut().enumerate() {
            *value = ((i + 1) * 10) as i32;
        }
    }
    println!("   Initialized array with values: [10, 20, 30, 40, 50]");

    println!("\n3. Accessing array elements:");

    // Test individual element access
    unsafe {
        assert_eq!(*caller_array.get(0).unwrap(), item_data[0]);
        assert_eq!(*caller_array.get(1).unwrap(), item_data[1]);
        assert_eq!(*caller_array.get(2).unwrap(), item_data[2]);
        assert!(caller_array.get(3).is_none()); // Out of bounds
    }
    println!("   ✓ Array element access works correctly");

    // Test mutable access
    unsafe {
        let mut item = *caller_array.get(1).unwrap();
        item.value = 25.0;
        *caller_array.as_mut_slice().unwrap().get_mut(1).unwrap() = item;
        assert_eq!(caller_array.get(1).unwrap().value, 25.0);
    }
    println!("   ✓ Mutable array access works correctly");
}

fn demonstrate_pointer_array_functions() {
    println!("\n\nPointer Array Functions Demonstration");
    println!("=====================================");

    println!("\n1. Creating pointer arrays:");

    // Create some pointers to simulate OPC item handles
    let ptr1 = CallerAllocatedPtr::<i32>::from_value(&100).unwrap();
    let ptr2 = CallerAllocatedPtr::<i32>::from_value(&200).unwrap();
    let ptr3 = CallerAllocatedPtr::<i32>::from_value(&300).unwrap();

    let ptrs = vec![ptr1.as_ptr(), ptr2.as_ptr(), ptr3.as_ptr()];
    let caller_ptr_array = CallerAllocatedPtrArray::from_ptr_slice(&ptrs).unwrap();
    println!(
        "   Created caller-allocated pointer array with {} pointers",
        caller_ptr_array.len()
    );

    // Verify the pointers were copied correctly
    unsafe {
        let slice = caller_ptr_array.as_slice().unwrap();
        assert_eq!(slice, ptrs.as_slice());
        println!("   ✓ Pointers were correctly copied to COM memory");
    }

    println!("\n2. Allocating uninitialized pointer arrays:");

    // Allocate memory for pointer array
    let mut uninit_ptr_array = CallerAllocatedPtrArray::<i32>::allocate(3).unwrap();
    println!("   Allocated uninitialized pointer array for 3 pointers");

    // Set pointers in the array
    unsafe {
        let test_ptr1 = std::ptr::null_mut::<i32>();
        let test_ptr2 = std::ptr::null_mut::<i32>();
        let test_ptr3 = std::ptr::null_mut::<i32>();

        assert!(uninit_ptr_array.set(0, test_ptr1));
        assert!(uninit_ptr_array.set(1, test_ptr2));
        assert!(uninit_ptr_array.set(2, test_ptr3));
        assert!(!uninit_ptr_array.set(3, test_ptr1)); // Out of bounds
    }
    println!("   ✓ Pointer array set operations work correctly");

    println!("\n3. Accessing pointer array elements:");

    // Test individual pointer access
    unsafe {
        assert_eq!(caller_ptr_array.get(0).unwrap(), ptrs[0]);
        assert_eq!(caller_ptr_array.get(1).unwrap(), ptrs[1]);
        assert_eq!(caller_ptr_array.get(2).unwrap(), ptrs[2]);
        assert!(caller_ptr_array.get(3).is_none()); // Out of bounds
    }
    println!("   ✓ Pointer array element access works correctly");
}

fn demonstrate_real_world_opc_scenario() {
    println!("\n\nReal-World OPC Array Scenario");
    println!("==============================");

    println!("\n1. Simulating OPC client preparing batch read request:");

    // Client prepares item IDs for batch read
    let item_ids = vec!["Item1", "Item2", "Item3", "Item4", "Item5"];
    let item_id_ptrs: Vec<*mut u16> = item_ids
        .iter()
        .map(|id| CallerAllocatedWString::from_str(id).unwrap().as_ptr())
        .collect();

    let item_id_array = CallerAllocatedPtrArray::from_ptr_slice(&item_id_ptrs).unwrap();
    println!(
        "   Client prepared {} item IDs for batch read",
        item_id_array.len()
    );

    println!("\n2. Simulating OPC server processing batch request:");

    // Server would process the request and return results
    println!("   Server would process the batch request");
    println!("   Server would allocate memory for results");

    println!("\n3. Simulating server returning batch results:");

    // Server would return arrays of results
    let result_values = vec![42.0, 84.0, 126.0, 168.0, 210.0];
    let result_qualities = vec![192u16, 192, 192, 192, 192];
    let result_timestamps = vec![
        1234567890u64,
        1234567891,
        1234567892,
        1234567893,
        1234567894,
    ];

    // In a real scenario, the server would allocate these arrays
    // and return pointers to the client
    println!("   Server would return:");
    println!("   - {} values: {:?}", result_values.len(), result_values);
    println!(
        "   - {} qualities: {:?}",
        result_qualities.len(),
        result_qualities
    );
    println!(
        "   - {} timestamps: {:?}",
        result_timestamps.len(),
        result_timestamps
    );

    println!("\n4. Client would receive and process batch results:");

    // Client would receive CalleeAllocatedArray types
    // and they would be automatically freed when they go out of scope
    println!("   Client would receive CalleeAllocatedArray types");
    println!("   Client would automatically free result memory");
    println!("   when the wrapper types go out of scope");

    println!("\n   ✓ Complete OPC batch operation cycle demonstrated");
}

fn main() {
    demonstrate_array_functions();
    demonstrate_pointer_array_functions();
    demonstrate_real_world_opc_scenario();

    println!("\n\n✅ Array functions demonstration completed!");
    println!("   These functions make COM array management much easier.");
    println!("   Perfect for batch operations in OPC Classic applications.");
}
