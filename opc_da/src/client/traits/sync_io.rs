use crate::utils::RemoteArray;
use windows::Win32::System::Variant::VARIANT;

/// Synchronous I/O functionality (OPC DA 1.0).
///
/// Provides methods for basic synchronous read/write operations
/// with direct server communication.
pub trait SyncIoTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO>;

    /// Reads values synchronously from items.
    ///
    /// # Arguments
    /// * `source` - Whether to read from cache or device
    /// * `server_handles` - Array of server item handles
    ///
    /// # Returns
    /// Tuple containing:
    /// - Array of item states (value, quality, timestamp)
    /// - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if server_handles is empty
    fn read(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<opc_da_bindings::tagOPCITEMSTATE>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut item_values = RemoteArray::new(len);
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Read(
                source,
                len,
                server_handles.as_ptr(),
                item_values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((item_values, errors))
    }

    /// Writes values synchronously to items.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles
    /// * `values` - Array of values to write
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn write(
        &self,
        server_handles: &[u32],
        values: &[VARIANT],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != values.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and values must have the same length",
            ));
        }

        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Write(
                len,
                server_handles.as_ptr(),
                values.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
