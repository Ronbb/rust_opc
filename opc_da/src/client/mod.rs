mod com;
mod def;
mod iter;

pub use com::*;
pub use def::*;
pub use iter::*;

pub struct Client {}

impl Client {
    pub fn get_servers(filter: ServerFilter) -> windows_core::Result<GuidIter> {
        let id = unsafe {
            windows::Win32::System::Com::CLSIDFromProgID(windows_core::w!("OPC.ServerList.1"))?
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

        let iter = unsafe {
            servers
                .EnumClassesOfCategories(
                    &filter
                        .available_versions
                        .iter()
                        .map(|v| v.to_guid())
                        .collect::<Vec<_>>(),
                    &filter
                        .requires_versions
                        .iter()
                        .map(|v| v.to_guid())
                        .collect::<Vec<_>>(),
                )
                .map_err(|e| {
                    windows_core::Error::new(e.code(), "Failed to enumerate server classes")
                })?
        };

        Ok(GuidIter::new(iter))
    }
}
