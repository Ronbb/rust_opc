use crate::client::{v1, v2, v3, ClientTrait as _};

use super::{GuidIterator, Server};

#[derive(Debug)]
pub enum Client {
    V1(v1::Client),
    V2(v2::Client),
    V3(v3::Client),
}

impl Client {
    pub fn v1() -> Self {
        Self::V1(v1::Client)
    }

    pub fn v2() -> Self {
        Self::V2(v2::Client)
    }

    pub fn v3() -> Self {
        Self::V3(v3::Client)
    }

    pub fn get_servers(&self) -> windows::core::Result<GuidIterator> {
        match self {
            Client::V1(client) => client.get_servers(),
            Client::V2(client) => client.get_servers(),
            Client::V3(client) => client.get_servers(),
        }
    }

    pub fn create_server(&self, class_id: windows::core::GUID) -> windows::core::Result<Server> {
        match self {
            Client::V1(client) => Ok(Server::V1(client.create_server(class_id)?)),
            Client::V2(client) => Ok(Server::V2(client.create_server(class_id)?)),
            Client::V3(client) => Ok(Server::V3(client.create_server(class_id)?)),
        }
    }
}
