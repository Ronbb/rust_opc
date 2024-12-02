mod def;
mod iter;

pub use def::*;
pub use iter::*;

pub struct Client {}

impl Client {
    pub fn get_servers(filter: ServerFilter) -> windows_core::Result<GuidIter> {
        unsafe {
            windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_MULTITHREADED,
            )
            .ok()?;
        };
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
            servers.EnumClassesOfCategories(
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
            )?
        };

        Ok(GuidIter(iter))
    }
}
