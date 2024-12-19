use windows::core::Interface as _;

use crate::client::{memory, ItemAttributeIterator};

/// Item management functionality.
///
/// Provides methods to manage OPC items within a group, including adding,
/// removing, and modifying item properties such as active state, client
/// handles, and data types.
pub trait ItemMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemMgt>;

    /// Adds items to the group.
    ///
    /// # Arguments
    /// * `items` - Array of item definitions containing item IDs and requested properties
    ///
    /// # Returns
    /// Tuple containing:
    /// - Array of item results with server handles and canonical data type
    /// - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if items array is empty
    fn add_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
    ) -> windows::core::Result<(
        memory::RemoteArray<opc_da_bindings::tagOPCITEMRESULT>,
        memory::RemoteArray<windows::core::HRESULT>,
    )> {
        if items.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "items cannot be empty",
            ));
        }

        let len = items.len().try_into()?;
        let mut results = memory::RemoteArray::new(len);
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?.AddItems(
                len,
                items.as_ptr(),
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((results, errors))
    }

    /// Validates item definitions without adding them to the group.
    ///
    /// # Arguments
    /// * `items` - Array of item definitions to validate
    /// * `blob_update` - Whether to validate blob update capability
    ///
    /// # Returns
    /// Tuple containing:
    /// - Array of item results with access rights and canonical data type
    /// - Array of per-item error codes
    fn validate_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
        blob_update: bool,
    ) -> windows::core::Result<(
        memory::RemoteArray<opc_da_bindings::tagOPCITEMRESULT>,
        memory::RemoteArray<windows::core::HRESULT>,
    )> {
        if items.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "items cannot be empty",
            ));
        }

        let len = items.len().try_into()?;
        let mut results = memory::RemoteArray::new(len);
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?.ValidateItems(
                len,
                items.as_ptr(),
                windows::Win32::Foundation::BOOL::from(blob_update),
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((results, errors))
    }

    /// Removes items from the group.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server handles for items to remove
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if server_handles is empty
    fn remove_items(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<memory::RemoteArray<windows::core::HRESULT>> {
        if server_handles.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?
                .RemoveItems(len, server_handles.as_ptr(), errors.as_mut_ptr())?;
        }

        Ok(errors)
    }

    /// Sets the active state of items.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server handles
    /// * `active` - True to activate items, false to deactivate
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if server_handles is empty
    fn set_active_state(
        &self,
        server_handles: &[u32],
        active: bool,
    ) -> windows::core::Result<memory::RemoteArray<windows::core::HRESULT>> {
        if server_handles.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?.SetActiveState(
                len,
                server_handles.as_ptr(),
                windows::Win32::Foundation::BOOL::from(active),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    /// Sets client handles for items.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server handles
    /// * `client_handles` - Array of new client handles
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn set_client_handles(
        &self,
        server_handles: &[u32],
        client_handles: &[u32],
    ) -> windows::core::Result<memory::RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != client_handles.len() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and client_handles must have the same length",
            ));
        }

        if server_handles.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?.SetClientHandles(
                len,
                server_handles.as_ptr(),
                client_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    /// Sets requested data types for items.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server handles
    /// * `requested_datatypes` - Array of VT_* data types
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn set_datatypes(
        &self,
        server_handles: &[u32],
        requested_datatypes: &[u16],
    ) -> windows::core::Result<memory::RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != requested_datatypes.len() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and requested_datatypes must have the same length",
            ));
        }

        if server_handles.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface()?.SetDatatypes(
                len,
                server_handles.as_ptr(),
                requested_datatypes.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    /// Creates an enumerator for item management.
    ///
    /// # Arguments
    /// * `id` - Interface ID specifying the type of enumerator
    ///
    /// # Returns
    /// Enumerator interface for iterating through items
    fn create_enumerator(&self) -> windows::core::Result<ItemAttributeIterator> {
        let enumerator = unsafe {
            self.interface()?
                .CreateEnumerator(&opc_da_bindings::IEnumOPCItemAttributes::IID)?
        };

        Ok(ItemAttributeIterator::new(enumerator.cast()?))
    }
}

// ...existing code...
