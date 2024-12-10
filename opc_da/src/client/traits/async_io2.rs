use crate::client::memory::RemoteArray;
use windows::core::VARIANT;

/// Asynchronous I/O functionality (OPC DA 2.0).
///
/// Provides enhanced asynchronous read/write operations without requiring
/// connection point callbacks. This trait extends the functionality of
/// AsyncIoTrait with improved error handling and control mechanisms.
pub trait AsyncIo2Trait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO2>;

    /// Initiates an asynchronous read operation.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles to read
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Tuple containing (cancel_id, error_array) where:
    /// - cancel_id: Identifier used to cancel the operation
    /// - error_array: Array of HRESULT values indicating per-item status
    fn read(
        &self,
        server_handles: &[u32],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        let len = server_handles.len().try_into()?;

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Read(
                len,
                server_handles.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

    /// Initiates an asynchronous write operation.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles to write
    /// * `values` - Array of VARIANT values to write
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Tuple containing (cancel_id, error_array) where:
    /// - cancel_id: Identifier used to cancel the operation
    /// - error_array: Array of HRESULT values indicating per-item status
    fn write(
        &self,
        server_handles: &[u32],
        values: &[VARIANT],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        let len = server_handles.len().try_into()?;

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Write(
                len,
                server_handles.as_ptr(),
                values.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

    /// Refreshes all active items from the specified source.
    ///
    /// # Arguments
    /// * `source` - Data source (cache or device)
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Cancel ID that can be used to cancel the operation
    fn refresh2(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        transaction_id: u32,
    ) -> windows::core::Result<u32> {
        unsafe { self.interface()?.Refresh2(source, transaction_id) }
    }

    /// Cancels a pending asynchronous operation.
    ///
    /// # Arguments
    /// * `cancel_id` - Cancel ID returned from read/write operations
    ///
    /// # Returns
    /// `Ok(())` if the operation was successfully canceled
    fn cancel2(&self, cancel_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.Cancel2(cancel_id) }
    }

    /// Enables or disables asynchronous I/O operations.
    ///
    /// # Arguments
    /// * `enable` - `true` to enable async operations, `false` to disable
    ///
    /// # Returns
    /// `Ok(())` if the enable state was successfully changed
    fn set_enable(&self, enable: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()?
                .SetEnable(windows::Win32::Foundation::BOOL::from(enable))
        }
    }

    /// Gets the current enable state of asynchronous I/O operations.
    ///
    /// # Returns
    /// `true` if async operations are enabled, `false` otherwise
    fn get_enable(&self) -> windows::core::Result<bool> {
        unsafe { self.interface()?.GetEnable().map(|v| v.as_bool()) }
    }
}
