use windows::Win32::System::Variant::VARIANT;

use crate::utils::RemoteArray;

/// Asynchronous I/O functionality (OPC DA 1.0).
///
/// Provides basic asynchronous read/write operations using connection point callbacks.
/// This is the original asynchronous interface defined in OPC DA 1.0.
pub trait AsyncIoTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO>;

    /// Reads values asynchronously from the server.
    ///
    /// # Arguments
    /// * `connection` - Connection point cookie for receiving callbacks
    /// * `source` - Specifies whether to read from cache or device
    /// * `server_handles` - Array of server item handles to read
    ///
    /// # Returns
    /// * `transaction_id` - Identifies this operation in callbacks
    /// * `errors` - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if server_handles is empty
    fn read(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut transaction_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Read(
                connection,
                source,
                len,
                server_handles.as_ptr(),
                &mut transaction_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((transaction_id, errors))
    }

    /// Writes values asynchronously to the server.
    ///
    /// # Arguments
    /// * `connection` - Connection point cookie for receiving callbacks
    /// * `server_handles` - Array of server item handles to write
    /// * `values` - Array of values to write
    ///
    /// # Returns
    /// * `transaction_id` - Identifies this operation in callbacks
    /// * `errors` - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn write(
        &self,
        connection: u32,
        server_handles: &[u32],
        values: &[VARIANT],
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
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

        let mut transaction_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Write(
                connection,
                len,
                server_handles.as_ptr(),
                values.as_ptr(),
                &mut transaction_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((transaction_id, errors))
    }

    /// Refreshes all active items asynchronously.
    ///
    /// # Arguments
    /// * `connection` - Connection point cookie for receiving callbacks
    /// * `source` - Specifies whether to refresh from cache or device
    ///
    /// # Returns
    /// Transaction ID for identifying the operation in callbacks
    fn refresh(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
    ) -> windows::core::Result<u32> {
        unsafe { self.interface()?.Refresh(connection, source) }
    }

    /// Cancels an outstanding asynchronous operation.
    ///
    /// # Arguments
    /// * `transaction_id` - ID of the operation to cancel
    ///
    /// # Returns
    /// Result indicating success or failure of cancel request
    fn cancel(&self, transaction_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.Cancel(transaction_id) }
    }
}
