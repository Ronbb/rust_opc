use windows_core::Interface as _;

use crate::{
    client::GuidIterator,
    def::{ClassContext, ServerInfo},
    utils::{IntoBridge, ToNative, TryToNative as _},
};

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

        let servers: opc_comn_bindings::IOPCServerList = unsafe {
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
    fn create_server(
        &self,
        class_id: windows::core::GUID,
        class_context: ClassContext,
    ) -> windows::core::Result<Server> {
        let server: opc_da_bindings::IOPCServer = unsafe {
            windows::Win32::System::Com::CoCreateInstance(
                &class_id,
                None,
                class_context.to_native(),
            )?
        };

        server.cast::<windows::core::IUnknown>()?.try_into()
    }

    fn create_server2(
        &self,
        class_id: windows::core::GUID,
        class_context: ClassContext,
        server_info: Option<ServerInfo>,
    ) -> windows::core::Result<Server> {
        let mut results = [windows::Win32::System::Com::MULTI_QI {
            pIID: &windows::core::IUnknown::IID,
            pItf: core::mem::ManuallyDrop::new(None),
            hr: windows::core::HRESULT(0),
        }];

        unsafe {
            windows::Win32::System::Com::CoCreateInstanceEx(
                &class_id,
                None,
                class_context.to_native(),
                match server_info {
                    Some(info) => Some(&info.into_bridge().try_to_native()?),
                    None => None,
                },
                &mut results,
            )?
        };

        if results[0].hr.is_err() {
            return Err(results[0].hr.into());
        }

        match results[0].pItf.as_ref() {
            Some(itf) => itf.cast::<windows::core::IUnknown>()?.try_into(),
            None => Err(windows::core::Error::from(
                windows::Win32::Foundation::E_POINTER,
            )),
        }
    }
}
