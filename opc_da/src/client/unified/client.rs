use windows::core::Interface;
use windows_core::IUnknown;

use super::{GuidIterator, Server};

#[derive(Debug)]
pub enum Client {
    V1,
    V2,
    V3,
}

impl Client {
    pub fn get_version_id(&self) -> windows::core::GUID {
        match self {
            Client::V1 => opc_da_bindings::CATID_OPCDAServer10::IID,
            Client::V2 => opc_da_bindings::CATID_OPCDAServer20::IID,
            Client::V3 => opc_da_bindings::CATID_OPCDAServer30::IID,
        }
    }

    pub fn get_servers(&self) -> windows::core::Result<GuidIterator> {
        let id = unsafe {
            windows::Win32::System::Com::CLSIDFromProgID(windows::core::w!("OPC.ServerList.1"))?
        };

        let servers: opc_da_bindings::IOPCServerList = unsafe {
            // TODO: Use CoCreateInstanceEx
            windows::Win32::System::Com::CoCreateInstance(
                &id,
                None,
                // TODO: Convert from filters
                windows::Win32::System::Com::CLSCTX_ALL,
            )?
        };

        let versions = [self.get_version_id()];

        let iter = unsafe {
            servers
                .EnumClassesOfCategories(&versions, &versions)
                .map_err(|e| {
                    windows::core::Error::new(e.code(), "Failed to enumerate server classes")
                })?
        };

        Ok(GuidIterator::new(iter))
    }

    pub fn create_server(&self, clsid: windows::core::GUID) -> windows::core::Result<Server> {
        let server: opc_da_bindings::IOPCServer = unsafe {
            windows::Win32::System::Com::CoCreateInstance(
                &clsid,
                None,
                windows::Win32::System::Com::CLSCTX_ALL,
            )?
        };

        let server: IUnknown = server.cast()?;

        match self {
            Client::V1 => Ok(Server::V1(server.try_into()?)),
            Client::V2 => Ok(Server::V2(server.try_into()?)),
            Client::V3 => Ok(Server::V3(server.try_into()?)),
        }
    }
}
