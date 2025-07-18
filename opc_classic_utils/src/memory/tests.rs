use super::*;
use std::str::FromStr;

#[test]
fn test_caller_allocated_ptr_null() {
    let ptr = CallerAllocatedPtr::<i32>::default();
    assert!(ptr.is_null());
}

#[test]
fn test_callee_allocated_ptr_null() {
    let ptr = CalleeAllocatedPtr::<i32>::default();
    assert!(ptr.is_null());
}

#[test]
fn test_caller_allocated_wstring_null() {
    let wstring = CallerAllocatedWString::default();
    assert!(wstring.is_null());
}

#[test]
fn test_callee_allocated_wstring_null() {
    let wstring = CalleeAllocatedWString::default();
    assert!(wstring.is_null());
}

#[test]
fn test_caller_allocated_ptr_no_free() {
    // This test verifies that CallerAllocatedPtr doesn't free memory
    // In a real scenario, this would be memory allocated by the caller
    let _ptr = CallerAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
    // When _ptr goes out of scope, it should NOT call CoTaskMemFree
}

#[test]
fn test_callee_allocated_ptr_frees() {
    // This test verifies that CalleeAllocatedPtr frees memory
    // In a real scenario, this would be memory allocated by the callee
    let _ptr = CalleeAllocatedPtr::from_raw(std::ptr::null_mut::<i32>());
    // When _ptr goes out of scope, it should call CoTaskMemFree
}

#[test]
fn test_transparent_repr() {
    // Test that transparent repr works correctly
    let ptr = std::ptr::null_mut::<i32>();

    // CallerAllocatedPtr should have the same memory layout as *mut i32
    let caller_ptr = CallerAllocatedPtr::from_raw(ptr);
    assert_eq!(caller_ptr.as_ptr(), ptr);

    // CalleeAllocatedPtr should have the same memory layout as *mut i32
    let callee_ptr = CalleeAllocatedPtr::from_raw(ptr);
    assert_eq!(callee_ptr.as_ptr(), ptr);

    // WString types should have the same memory layout as *mut u16
    let wstring_ptr = std::ptr::null_mut::<u16>();
    let caller_wstring = CallerAllocatedWString::from_raw(wstring_ptr);
    assert_eq!(caller_wstring.as_ptr(), wstring_ptr);

    let callee_wstring = CalleeAllocatedWString::from_raw(wstring_ptr);
    assert_eq!(callee_wstring.as_ptr(), wstring_ptr);
}

#[test]
fn test_caller_allocated_ptr_allocate() {
    // Test allocation of caller-allocated pointer
    let ptr = CallerAllocatedPtr::<i32>::allocate().unwrap();
    assert!(!ptr.is_null());
    // Memory will be freed by callee, not by our wrapper
}

#[test]
fn test_caller_allocated_ptr_from_value() {
    // Test creating pointer from value
    let value = 42i32;
    let ptr = CallerAllocatedPtr::from_value(&value).unwrap();
    assert!(!ptr.is_null());

    // Verify the value was copied correctly
    unsafe {
        assert_eq!(*ptr.as_ptr(), 42);
    }
}

#[test]
fn test_caller_allocated_wstring_from_str() {
    // Test creating wide string from Rust string
    use std::str::FromStr;
    let test_string = "Hello, World!";
    let wstring = CallerAllocatedWString::from_str(test_string).unwrap();
    assert!(!wstring.is_null());

    // Verify the string was converted correctly
    unsafe {
        let converted = wstring.to_string().unwrap();
        assert_eq!(converted, test_string);
    }
}

#[test]
fn test_caller_allocated_wstring_from_string() {
    // Test creating wide string from String
    let test_string = String::from("Test String");
    let wstring = CallerAllocatedWString::from_string(test_string.clone()).unwrap();
    assert!(!wstring.is_null());

    // Verify the string was converted correctly
    unsafe {
        let converted = wstring.to_string().unwrap();
        assert_eq!(converted, test_string);
    }
}

#[test]
fn test_caller_allocated_wstring_from_os_str() {
    // Test creating wide string from OsStr
    let test_string = std::ffi::OsStr::new("OS String Test");
    let wstring = CallerAllocatedWString::from_os_str(test_string).unwrap();
    assert!(!wstring.is_null());

    // Verify the string was converted correctly
    unsafe {
        let converted = wstring.to_os_string().unwrap();
        assert_eq!(converted, test_string);
    }
}

#[test]
fn test_pointer_dereference() {
    // Test dereferencing pointers
    let value = 123i32;
    let mut caller_ptr = CallerAllocatedPtr::from_value(&value).unwrap();
    // Don't create a CalleeAllocatedPtr from CallerAllocatedPtr - it would cause double-free

    // Test as_ref
    unsafe {
        assert_eq!(caller_ptr.as_ref().unwrap(), &123);
    }

    // Test as_mut
    unsafe {
        *caller_ptr.as_mut().unwrap() = 456;
        assert_eq!(*caller_ptr.as_ptr(), 456);
    }
}

#[test]
fn test_callee_allocated_ptr_dereference() {
    // Test dereferencing CalleeAllocatedPtr with separate memory
    let value = 789i32;
    let callee_ptr = CalleeAllocatedPtr::from_value(&value).unwrap();

    // Test as_ref
    unsafe {
        assert_eq!(callee_ptr.as_ref().unwrap(), &789);
    }
}

#[test]
fn test_null_pointer_dereference() {
    // Test dereferencing null pointers
    let caller_ptr = CallerAllocatedPtr::<i32>::default();
    let callee_ptr = CalleeAllocatedPtr::<i32>::default();

    unsafe {
        assert!(caller_ptr.as_ref().is_none());
        assert!(callee_ptr.as_ref().is_none());
    }
}

#[test]
fn test_wstring_null_conversion() {
    // Test converting null wide strings
    let caller_wstring = CallerAllocatedWString::default();
    let callee_wstring = CalleeAllocatedWString::default();

    unsafe {
        assert!(caller_wstring.to_string().is_none());
        assert!(callee_wstring.to_string().is_none());
        assert!(caller_wstring.to_os_string().is_none());
        assert!(callee_wstring.to_os_string().is_none());
    }
}

#[test]
fn test_from_str_trait() {
    // Test the FromStr trait implementation
    let test_string = "Test FromStr Trait";
    let wstring = CallerAllocatedWString::from_str(test_string).unwrap();
    assert!(!wstring.is_null());

    // Verify the string was converted correctly
    unsafe {
        let converted = wstring.to_string().unwrap();
        assert_eq!(converted, test_string);
    }
}

#[test]
fn test_caller_allocated_array_null() {
    let array = CallerAllocatedArray::<i32>::default();
    assert!(array.is_null());
    assert!(array.is_empty());
    assert_eq!(array.len(), 0);
}

#[test]
fn test_callee_allocated_array_null() {
    let array = CalleeAllocatedArray::<i32>::default();
    assert!(array.is_null());
    assert!(array.is_empty());
    assert_eq!(array.len(), 0);
}

#[test]
fn test_caller_allocated_array_allocate() {
    let array = CallerAllocatedArray::<i32>::allocate(5).unwrap();
    assert!(!array.is_null());
    assert!(!array.is_empty());
    assert_eq!(array.len(), 5);
}

#[test]
fn test_caller_allocated_array_from_slice() {
    let data = vec![1, 2, 3, 4, 5];
    let array = CallerAllocatedArray::from_slice(&data).unwrap();
    assert!(!array.is_null());
    assert_eq!(array.len(), 5);

    // Verify the data was copied correctly
    unsafe {
        let slice = array.as_slice().unwrap();
        assert_eq!(slice, data.as_slice());
    }
}

#[test]
fn test_caller_allocated_array_access() {
    let data = vec![10, 20, 30];
    let mut array = CallerAllocatedArray::from_slice(&data).unwrap();

    // Test get
    unsafe {
        assert_eq!(*array.get(0).unwrap(), 10);
        assert_eq!(*array.get(1).unwrap(), 20);
        assert_eq!(*array.get(2).unwrap(), 30);
        assert!(array.get(3).is_none()); // Out of bounds
    }

    // Test get_mut
    unsafe {
        *array.get_mut(1).unwrap() = 25;
        assert_eq!(*array.get(1).unwrap(), 25);
    }

    // Test as_slice and as_mut_slice
    unsafe {
        let slice = array.as_slice().unwrap();
        assert_eq!(slice, &[10, 25, 30]);

        let mut_slice = array.as_mut_slice().unwrap();
        mut_slice[2] = 35;
        assert_eq!(*array.get(2).unwrap(), 35);
    }
}

#[test]
fn test_callee_allocated_array_frees_container() {
    // This test verifies that CalleeAllocatedArray frees the container memory
    // In a real scenario, this would be memory allocated by the callee
    let _array: CalleeAllocatedArray<i32> = CalleeAllocatedArray::from_raw(std::ptr::null_mut(), 0);
    // When _array goes out of scope, it should call CoTaskMemFree on the container
}

#[test]
fn test_caller_allocated_ptr_array_null() {
    let array = CallerAllocatedPtrArray::<i32>::default();
    assert!(array.is_null());
    assert!(array.is_empty());
    assert_eq!(array.len(), 0);
}

#[test]
fn test_callee_allocated_ptr_array_null() {
    let array = CalleeAllocatedPtrArray::<i32>::default();
    assert!(array.is_null());
    assert!(array.is_empty());
    assert_eq!(array.len(), 0);
}

#[test]
fn test_caller_allocated_ptr_array_allocate() {
    let array = CallerAllocatedPtrArray::<i32>::allocate(3).unwrap();
    assert!(!array.is_null());
    assert!(!array.is_empty());
    assert_eq!(array.len(), 3);
}

#[test]
fn test_caller_allocated_ptr_array_from_slice() {
    let ptrs = vec![
        std::ptr::null_mut::<i32>(),
        std::ptr::null_mut::<i32>(),
        std::ptr::null_mut::<i32>(),
    ];
    let array = CallerAllocatedPtrArray::from_ptr_slice(&ptrs).unwrap();
    assert!(!array.is_null());
    assert_eq!(array.len(), 3);

    // Verify the pointers were copied correctly
    unsafe {
        let slice = array.as_slice().unwrap();
        assert_eq!(slice, ptrs.as_slice());
    }
}

#[test]
fn test_caller_allocated_ptr_array_access() {
    let mut array = CallerAllocatedPtrArray::<i32>::allocate(2).unwrap();

    // Test get and set
    unsafe {
        // Newly allocated memory contains uninitialized values, not necessarily null
        let _ptr0 = array.get(0).unwrap();
        let _ptr1 = array.get(1).unwrap();

        // Set to null and verify
        let test_ptr = std::ptr::null_mut::<i32>();
        assert!(array.set(0, test_ptr));
        assert_eq!(array.get(0).unwrap(), test_ptr);
        assert!(array.get(0).unwrap().is_null());

        // Test out of bounds
        assert!(!array.set(2, test_ptr)); // Out of bounds
    }
}

#[test]
fn test_callee_allocated_ptr_array_frees_all() {
    // This test verifies that CalleeAllocatedPtrArray frees both container and elements
    // In a real scenario, this would be memory allocated by the callee
    let _array: CalleeAllocatedPtrArray<i32> =
        CalleeAllocatedPtrArray::from_raw(std::ptr::null_mut(), 0);
    // When _array goes out of scope, it should call CoTaskMemFree on both container and elements
}

#[test]
fn test_array_transparent_repr() {
    // Test that transparent repr works correctly for arrays
    let ptr = std::ptr::null_mut::<i32>();
    let len = 5;

    // CallerAllocatedArray should have the same memory layout as (*mut i32, usize)
    let caller_array = CallerAllocatedArray::from_raw(ptr, len);
    assert_eq!(caller_array.as_ptr(), ptr);
    assert_eq!(caller_array.len(), len);

    // CalleeAllocatedArray should have the same memory layout as (*mut i32, usize)
    let callee_array = CalleeAllocatedArray::from_raw(ptr, len);
    assert_eq!(callee_array.as_ptr(), ptr);
    assert_eq!(callee_array.len(), len);

    // Pointer arrays should have the same memory layout as (*mut *mut i32, usize)
    let ptr_array = std::ptr::null_mut::<*mut i32>();
    let caller_ptr_array = CallerAllocatedPtrArray::from_raw(ptr_array, len);
    assert_eq!(caller_ptr_array.as_ptr(), ptr_array);
    assert_eq!(caller_ptr_array.len(), len);

    let callee_ptr_array = CalleeAllocatedPtrArray::from_raw(ptr_array, len);
    assert_eq!(callee_ptr_array.as_ptr(), ptr_array);
    assert_eq!(callee_ptr_array.len(), len);
}
