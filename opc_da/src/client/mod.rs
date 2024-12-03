mod def;
mod iter;

pub use def::*;
pub use iter::*;

pub struct Client {}

impl Client {
    pub fn ensure_com() -> windows_core::Result<()> {
        static COM_INIT: std::sync::Once = std::sync::Once::new();
        static mut COM_RESULT: Option<windows_core::HRESULT> = None;

        unsafe {
            COM_INIT.call_once(|| {
                COM_RESULT = Some(windows::Win32::System::Com::CoInitializeEx(
                    None,
                    windows::Win32::System::Com::COINIT_MULTITHREADED,
                ));
            });
            COM_RESULT.unwrap_or(windows::Win32::Foundation::S_OK).ok()
        }
    }

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

impl Drop for Client {
    fn drop(&mut self) {
        unsafe {
            windows::Win32::System::Com::CoUninitialize();
        }
    }
}
