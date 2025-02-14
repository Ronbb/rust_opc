use unified::{Guard, Server};

use crate::utils::LocalPointer;

use super::*;

#[test]
fn test_client() {
    let client = Guard::new(unified::Client::v2()).expect("Failed to create guard");

    let servers = client
        .get_servers()
        .expect("Failed to get servers")
        .collect::<Vec<_>>();

    if servers.is_empty() {
        panic!("No servers found");
    }

    let server_id = servers
        .first()
        .cloned()
        .expect("No server found")
        .expect("Failed to get server id");

    println!("Server ID: {:?}", server_id);

    let server = client
        .create_server(server_id)
        .expect("Failed to create server");

    let server = match server {
        Server::V2(server) => server,
        _ => panic!("Expected V2 server"),
    };

    let branch = StringIterator::new(
        server
            .browse_opc_item_ids(opc_da_bindings::OPC_BRANCH, Some(""), 0, 0)
            .expect("Failed to browse items"),
    )
    .take(1)
    .collect::<Result<Vec<_>, _>>()
    .expect("No names found")
    .pop()
    .expect("No branch found");

    println!("Branch: {:?}", branch);

    server
        .change_browse_position(opc_da_bindings::OPC_BROWSE_TO, &branch)
        .expect("Failed to change browse position");

    let leaf = StringIterator::new(
        server
            .browse_opc_item_ids(opc_da_bindings::OPC_FLAT, Some(""), 0, 0)
            .expect("Failed to browse items"),
    )
    .take(1)
    .collect::<Result<Vec<_>, _>>()
    .expect("No names found")
    .pop()
    .expect("No leaf found");

    println!("Leaf: {:?}", leaf);

    let name = server.get_item_id(&leaf).expect("Failed to get item id");

    println!("Item name: {:?}", name);

    let mut group_server_handle = 0u32;
    let mut revised_percent_deadband = 0u32;
    let group = server
        .add_group(
            "test",
            true,
            0,
            0,
            0,
            0,
            0.0,
            &mut revised_percent_deadband,
            &mut group_server_handle,
        )
        .expect("Failed to add group");

    let name = LocalPointer::from(&name);
    let (results, errors) = group
        .add_items(&[opc_da_bindings::tagOPCITEMDEF {
            szAccessPath: windows::core::PWSTR::null(),
            szItemID: name.as_pwstr(),
            bActive: true.into(),
            hClient: 0,
            dwBlobSize: 0,
            pBlob: std::ptr::null_mut(),
            vtRequestedDataType: 0,
            wReserved: 0,
        }])
        .expect("Failed to add items");

    assert_eq!(errors.len(), 1, "Expected exactly one error result");
    let error = errors.as_slice().first().expect("Error array is empty");
    assert!(error.is_ok(), "Unexpected error: {:?}", error);

    assert_eq!(results.len(), 1, "Expected exactly one result");
    let server_handle = results
        .as_slice()
        .first()
        .unwrap_or_else(|| panic!("Expected 1 result, got {}", results.len()))
        .hServer;

    let (states, errors) =
        SyncIoTrait::read(&group, opc_da_bindings::OPC_DS_CACHE, &[server_handle])
            .expect("Failed to read");

    if errors.len() != 1 {
        panic!("Expected 1 error, got {}", errors.len());
    }

    let error = errors.as_slice().first().unwrap();
    if error.is_err() {
        panic!("Error, got {:?}", error);
    }

    if states.len() != 1 {
        panic!("Expected 1 state, got {}", states.len());
    }

    let state = states.as_slice().first().unwrap();
    println!("State: {:?}", state.vDataValue);

    let cloned_value = state.vDataValue.clone();
    SyncIoTrait::write(&group, &[server_handle], &[cloned_value]).expect("Failed to write");
}
