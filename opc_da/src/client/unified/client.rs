use windows::core::Interface;

use crate::def;

use super::{GuidIterator, Server};

#[derive(Debug)]
pub struct Client {
    /// Marker to ensure `Client` is not `Send` and not `Sync`.
    _marker: std::marker::PhantomData<*const ()>,
}

impl Client {
    pub fn new() -> windows::core::Result<Self> {
        Self::initialize().ok()?;
        Ok(Self {
            _marker: std::marker::PhantomData,
        })
    }
}

impl Drop for Client {
    fn drop(&mut self) {
        unsafe { Self::uninitialize() };
    }
}

impl Client {
    /// Ensures COM is initialized for the current thread.
    ///
    /// # Returns
    /// Returns the HRESULT of the COM initialization.
    ///
    /// # Thread Safety
    /// COM initialization is performed with COINIT_MULTITHREADED flag.
    ///
    /// # Note
    /// Callers should check the returned HRESULT for initialization failures.
    pub(crate) fn initialize() -> windows::core::HRESULT {
        unsafe {
            windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_MULTITHREADED,
            )
        }
    }

    /// Uninitializes COM for the current thread.
    ///
    /// # Safety
    /// This method should be called when the thread is shutting down
    /// and no more COM calls will be made.
    pub(crate) unsafe fn uninitialize() {
        windows::Win32::System::Com::CoUninitialize();
    }

    pub fn get_servers(&self, filter: def::ServerFilter) -> windows::core::Result<GuidIterator> {
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

        server.cast::<windows::core::IUnknown>()?.try_into()
    }
}
