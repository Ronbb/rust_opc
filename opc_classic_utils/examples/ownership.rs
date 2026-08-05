use opc_classic_utils::{CoTaskMemArrayBuilder, DropElements, WideCString};

fn main() -> windows_core::Result<()> {
    // COM input parameters borrow ordinary Rust-owned memory.
    let item_id =
        WideCString::try_from("Channel.Device.Tag").expect("the item id must not contain NUL");
    println!("input pointer: {:?}", item_id.as_pcwstr());

    // COM outputs use the task allocator and an explicit element cleanup policy.
    let mut values = CoTaskMemArrayBuilder::new(3, DropElements)?;
    values.push(String::from("one")).unwrap();
    values.push(String::from("two")).unwrap();
    values.push(String::from("three")).unwrap();
    let values = values.finish()?;
    println!("values: {:?}", values.as_slice());
    Ok(())
}
