use crate::def::*;

use super::*;

#[test]
fn test_unified() {
    let client = Guard::new(Client::v2()).expect("Failed to create client guard");
    let mut servers = client.get_servers().expect("Failed to get servers");
    let server_id = servers
        .next()
        .expect("No servers found")
        .expect("Failed to get server id");

    let server = client
        .create_server(server_id)
        .expect("Failed to create server");

    let group_state = GroupState::default();
    let _ = server.add_group(group_state).expect("Failed to add group");
}
