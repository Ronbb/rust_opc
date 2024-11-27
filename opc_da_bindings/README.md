# OPC DA Bindings

Please see docs on [docs.rs](https://docs.rs/opc_da_bindings/).

## Example

### Enumerate server list

```rust
use opc_da_bindings::{
    CATID_OPCDAServer10, CATID_OPCDAServer20, CATID_OPCDAServer30, IOPCServerList,
};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, ProgIDFromCLSID, CLSCTX_ALL,
    COINIT_MULTITHREADED,
};
use windows_core::{w, Interface, GUID};

pub fn get_servers() -> Vec<String> {
    unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).unwrap() };
    let id = unsafe { CLSIDFromProgID(w!("OPC.ServerList.1")).unwrap() };

    let servers: IOPCServerList = unsafe { CoCreateInstance(&id, None, CLSCTX_ALL).unwrap() };

    let enumer = unsafe {
        servers
            .EnumClassesOfCategories(
                &[
                    CATID_OPCDAServer10::IID,
                    CATID_OPCDAServer20::IID,
                    CATID_OPCDAServer30::IID,
                ],
                &[],
            )
            .unwrap()
    };

    let mut results = vec![];

    loop {
        let mut server_id = vec![GUID::zeroed(); 8];
        let mut count = 0;

        unsafe { enumer.Next(&mut server_id, Some(&mut count)).unwrap() };

        if count == 0 {
            break;
        }

        for i in 0..count {
            let id = server_id[i as usize];
            unsafe {
                let id = ProgIDFromCLSID(&id).unwrap();
                println!("ProgID: {}", id.to_string().unwrap());

                results.push(id.to_string().unwrap());
            }
        }
    }

    results
}
```
