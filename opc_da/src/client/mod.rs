mod actor;
mod com;
mod def;
mod iter;

pub use actor::*;
pub use def::*;
pub use iter::*;

#[derive(Debug)]
pub struct Client {
    /// Marker to ensure `Client` is not `Send` and not `Sync`.
    _marker: std::marker::PhantomData<*const ()>,
}

impl Client {
    pub fn new() -> windows_core::Result<Self> {
        Self::initialize().ok()?;
        Ok(Self {
            _marker: std::marker::PhantomData,
        })
    }

    pub fn start() -> windows_core::Result<ClientAsync> {
        Client::new().map(Into::into)
    }

    pub fn get_servers(&self, filter: ServerFilter) -> windows_core::Result<GuidIter> {
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
        unsafe { Self::uninitialize() };
    }
}
