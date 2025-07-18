//! Convenience functions demonstration for opc_classic_utils
//!
//! This example shows how to use the convenience functions for
//! creating COM-allocated memory from Rust types and string conversions.

use opc_classic_utils::memory::{CallerAllocatedPtr, CallerAllocatedWString};

#[derive(Debug, Clone, Copy, PartialEq)]
struct TestStruct {
    id: i32,
    value: f64,
}

fn demonstrate_pointer_convenience_functions() {
    println!("Pointer Convenience Functions");
    println!("============================");

    println!("\n1. Creating pointers from values:");

    // Create a pointer from a simple value
    let int_value = 42;
    let int_ptr = CallerAllocatedPtr::from_value(&int_value).unwrap();
    println!("   Created pointer from int: {:?}", int_value);

    // Create a pointer from a struct
    let struct_value = TestStruct {
        id: 1,
        value: std::f64::consts::PI,
    };
    let struct_ptr = CallerAllocatedPtr::from_value(&struct_value).unwrap();
    println!("   Created pointer from struct: {:?}", struct_value);

    // Verify the values were copied correctly
    unsafe {
        assert_eq!(*int_ptr.as_ptr(), 42);
        assert_eq!(*struct_ptr.as_ptr(), struct_value);
    }

    println!("   ✓ Values were correctly copied to COM memory");

    println!("\n2. Allocating uninitialized memory:");

    // Allocate memory for later initialization
    let uninit_ptr = CallerAllocatedPtr::<i32>::allocate().unwrap();
    println!("   Allocated uninitialized memory for i32");

    // Initialize the memory
    unsafe {
        *uninit_ptr.as_ptr() = 100;
    }
    println!("   Initialized memory with value: 100");

    println!("\n3. Dereferencing pointers:");

    let value = 123;
    let mut ptr = CallerAllocatedPtr::from_value(&value).unwrap();

    // Use as_ref for read-only access
    unsafe {
        let ref_value = ptr.as_ref().unwrap();
        println!("   Read value through as_ref: {}", ref_value);
    }

    // Use as_mut for mutable access
    unsafe {
        let mut_value = ptr.as_mut().unwrap();
        *mut_value = 456;
        println!("   Modified value through as_mut: {}", mut_value);
    }

    // Verify the change
    unsafe {
        assert_eq!(*ptr.as_ptr(), 456);
    }
    println!("   ✓ Value was successfully modified");
}

fn demonstrate_string_convenience_functions() {
    println!("\n\nString Convenience Functions");
    println!("============================");

    println!("\n1. Creating wide strings from Rust strings:");

    // From string slice
    use std::str::FromStr;
    let str_slice = "Hello, World!";
    let wstring1 = CallerAllocatedWString::from_str(str_slice).unwrap();
    println!("   Created from &str: '{}'", str_slice);

    // From String
    let owned_string = String::from("Owned String");
    let wstring2 = CallerAllocatedWString::from_string(owned_string.clone()).unwrap();
    println!("   Created from String: '{}'", owned_string);

    // From OsStr
    use std::ffi::OsStr;
    let os_string = OsStr::new("OS String");
    let wstring3 = CallerAllocatedWString::from_os_str(os_string).unwrap();
    println!("   Created from OsStr: '{:?}'", os_string);

    println!("\n2. Converting wide strings back to Rust strings:");

    // Convert back to String
    unsafe {
        let converted1 = wstring1.to_string().unwrap();
        println!("   Converted back to String: '{}'", converted1);
        assert_eq!(converted1, str_slice);

        let converted2 = wstring2.to_string().unwrap();
        println!("   Converted back to String: '{}'", converted2);
        assert_eq!(converted2, owned_string);

        let converted3 = wstring3.to_os_string().unwrap();
        println!("   Converted back to OsString: '{:?}'", converted3);
        assert_eq!(converted3, os_string);
    }

    println!("   ✓ All string conversions work correctly");

    println!("\n3. Handling null strings:");

    let null_wstring = CallerAllocatedWString::default();
    unsafe {
        let result = null_wstring.to_string();
        assert!(result.is_none());
        println!("   Null string correctly returns None");
    }
}

fn demonstrate_real_world_scenario() {
    use std::str::FromStr;

    println!("\n\nReal-World OPC Scenario");
    println!("=======================");

    println!("\n1. Simulating OPC client preparing input parameters:");

    // Client allocates memory for input parameters
    let server_name = CallerAllocatedWString::from_str("OPC.Server.1").unwrap();
    let item_count = CallerAllocatedPtr::from_value(&5i32).unwrap();

    println!("   Client allocated memory for:");
    println!("   - Server name: 'OPC.Server.1'");
    println!("   - Item count: 5");

    println!("\n2. Simulating OPC server processing:");

    // Server would use these parameters
    unsafe {
        let name = server_name.to_string().unwrap();
        let count = *item_count.as_ptr();
        println!("   Server received:");
        println!("   - Server name: '{}'", name);
        println!("   - Item count: {}", count);
    }

    println!("\n3. Simulating server returning output parameters:");

    // Server would allocate memory for output
    let result_string = "Operation completed successfully";
    let result_code = 0i32;

    // In a real scenario, the server would allocate this memory
    // and return pointers to the client
    println!("   Server would return:");
    println!("   - Result message: '{}'", result_string);
    println!("   - Result code: {}", result_code);

    println!("\n4. Client would receive and process output:");

    // Client would receive CalleeAllocatedPtr/CalleeAllocatedWString
    // and they would be automatically freed when they go out of scope
    println!("   Client would automatically free output memory");
    println!("   when the wrapper types go out of scope");

    println!("\n   ✓ Complete OPC memory management cycle demonstrated");
}

fn main() {
    demonstrate_pointer_convenience_functions();
    demonstrate_string_convenience_functions();
    demonstrate_real_world_scenario();

    println!("\n\n✅ Convenience functions demonstration completed!");
    println!("   These functions make COM memory management much easier.");
}
