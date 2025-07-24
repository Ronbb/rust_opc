//! Example demonstrating the usage of convenience macros for OPC Classic memory management
//!
//! This example shows how to use the new macros:
//! - `write_caller_allocated_ptr!` - for writing to caller-allocated pointers
//! - `write_caller_allocated_array!` - for writing to caller-allocated arrays
//! - `alloc_callee_wstring!` - for allocating callee-allocated wide strings

use opc_classic_utils::{
    alloc_callee_wstring, write_caller_allocated_array, write_caller_allocated_ptr,
};

fn main() -> windows::core::Result<()> {
    println!("=== OPC Classic Utils Macro Usage Examples ===\n");

    // Example 1: Using write_caller_allocated_ptr! macro
    println!("1. Using write_caller_allocated_ptr! macro:");
    {
        let mut count: u32 = 0;
        let count_ptr = &mut count as *mut u32;

        // Using the macro to write to a caller-allocated pointer
        write_caller_allocated_ptr!(count_ptr, 42u32)?;
        println!("   Written value: {}", count);
    }

    // Example 2: Using write_caller_allocated_array! macro
    println!("\n2. Using write_caller_allocated_array! macro:");
    {
        let mut array_ptr: *mut u32 = std::ptr::null_mut();

        // Using the macro to write an array to a caller-allocated pointer
        let data = vec![1u32, 2u32, 3u32, 4u32, 5u32];
        write_caller_allocated_array!(&mut array_ptr, &data)?;

        println!("   Array pointer: {:?}", array_ptr);
        println!("   Array is not null: {}", !array_ptr.is_null());
    }

    // Example 3: Using alloc_callee_wstring! macro
    println!("\n3. Using alloc_callee_wstring! macro:");
    {
        let error_message = "This is an error message";

        // Using the macro to allocate a callee-allocated wide string
        let wide_string_ptr = alloc_callee_wstring!(error_message)?;

        println!("   Wide string pointer: {:?}", wide_string_ptr);
    }

    // Example 4: Simulating OPC Common interface usage
    println!("\n4. Simulating OPC Common interface usage:");
    {
        // Simulate QueryAvailableLocaleIDs method
        let mut count: u32 = 0;
        let mut locale_ids_ptr: *mut u32 = std::ptr::null_mut();

        // Simulate available locale IDs
        let available_locale_ids = vec![0x0409u32, 0x0410u32, 0x0411u32]; // EN-US, IT, JA

        // Write count using macro
        write_caller_allocated_ptr!(&mut count, available_locale_ids.len() as u32)?;

        // Write array using macro
        write_caller_allocated_array!(&mut locale_ids_ptr, &available_locale_ids)?;

        println!("   Available locale count: {}", count);
        println!("   Locale IDs pointer: {:?}", locale_ids_ptr);
    }

    // Example 5: Error handling with macros
    println!("\n5. Error handling with macros:");
    {
        // Simulate getting an error string from an HRESULT
        let hresult = windows::Win32::Foundation::E_POINTER;
        let error_string_ptr = alloc_callee_wstring!("Pointer is invalid")?;

        println!("   HRESULT: {:?}", hresult);
        println!("   Error string pointer: {:?}", error_string_ptr);
    }

    println!("\n=== All examples completed successfully! ===");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_caller_allocated_ptr_macro() {
        let mut value: u32 = 0;
        let ptr = &mut value as *mut u32;

        let result = write_caller_allocated_ptr!(ptr, 123u32);
        assert!(result.is_ok());
        assert_eq!(value, 123);
    }

    #[test]
    fn test_write_caller_allocated_array_macro() {
        let mut array_ptr: *mut u32 = std::ptr::null_mut();
        let data = vec![1u32, 2u32, 3u32];

        let result = write_caller_allocated_array!(&mut array_ptr, &data);
        assert!(result.is_ok());
        assert!(!array_ptr.is_null());
    }

    #[test]
    fn test_alloc_callee_wstring_macro() {
        let result = alloc_callee_wstring!("test string");
        assert!(result.is_ok());
    }
}
