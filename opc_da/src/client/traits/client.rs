use windows_core::Interface as _;

use crate::client::GuidIterator;

/// Trait defining client functionality for OPC Data Access servers.
pub trait ClientTrait<Server: TryFrom<windows::core::IUnknown, Error = windows::core::Error>> {
    /// GUID of the catalog used to enumerate servers.
    const CATALOG_ID: windows::core::GUID;

    /// Retrieves an iterator over available server GUIDs.
    ///
    /// # Returns
    ///
    /// A `Result` containing a `GuidIterator` over server GUIDs, or an error if the operation fails.
    fn get_servers(&self) -> windows::core::Result<GuidIterator> {
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

        let versions = [Self::CATALOG_ID];

        let iter = unsafe {
            servers
                .EnumClassesOfCategories(&versions, &versions)
                .map_err(|e| {
                    windows::core::Error::new(e.code(), "Failed to enumerate server classes")
                })?
        };

        Ok(GuidIterator::new(iter))
    }

    /// Creates a server instance from the specified class ID.
    ///
    /// # Parameters
    ///
    /// - `class_id`: The GUID of the server class to instantiate.
    ///
    /// # Returns
    ///
    /// A `Result` containing the server instance, or an error if creation fails.
    fn create_server(&self, class_id: windows::core::GUID) -> windows::core::Result<Server> {
        let server: opc_da_bindings::IOPCServer = unsafe {
            windows::Win32::System::Com::CoCreateInstance(
                &class_id,
                None,
                windows::Win32::System::Com::CLSCTX_ALL,
            )?
        };

        server.cast::<windows::core::IUnknown>()?.try_into()
    }
}
