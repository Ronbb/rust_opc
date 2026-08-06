use opc_classic_utils::{CoTaskMemArrayBuilder, Error, FreePwstrElements, OwnedPwstr, WideCString};

fn main() -> opc_classic_utils::Result<()> {
    // COM input parameters borrow ordinary Rust-owned memory.
    let item_id =
        WideCString::try_from("Channel.Device.Tag").expect("the item id must not contain NUL");
    println!("input pointer: {:?}", item_id.as_ptr());

    // COM outputs use the task allocator and an explicit element cleanup policy.
    let mut values = CoTaskMemArrayBuilder::new(3, FreePwstrElements)?;
    for text in ["one", "two", "three"] {
        let raw = OwnedPwstr::new(text)?.into_raw();
        if let Err(raw) = values.push(raw) {
            drop(unsafe { OwnedPwstr::from_raw(raw) });
            return Err(Error::unexpected("example output array is full"));
        }
    }
    let values = values.finish()?;
    println!("{} task-allocated strings", values.len());
    Ok(())
}
