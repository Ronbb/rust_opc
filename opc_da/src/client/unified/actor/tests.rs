use crate::client::unified::{create_runtime, Client};

use super::ClientActor;

#[test]
fn test_actor() {
    actix::System::with_tokio_rt(create_runtime).block_on(async {
        let client = ClientActor::new(Client::v2()).expect("Failed to create client actor");
        let servers = client.get_servers().await.expect("Failed to get servers");

        assert!(!servers.is_empty());
    });
}
