use crate::client::memory::RemoteArray;

/// Asynchronous I/O functionality (OPC DA 3.0).
///
/// Provides methods for enhanced asynchronous read/write operations with
/// quality and timestamp information.
pub trait AsyncIo3Trait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO3>;

    /// Reads values with maximum age constraint.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles
    /// * `max_age` - Maximum age constraints for each item (milliseconds)
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Tuple containing cancel ID and array of per-item error codes
    fn read_max_age(
        &self,
        server_handles: &[u32],
        max_age: &[u32],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.len() != max_age.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and max_age must have the same length",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.ReadMaxAge(
                len,
                server_handles.as_ptr(),
                max_age.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

    /// Writes values with quality and timestamp information.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles
    /// * `values` - Array of values with quality and timestamp (VQT)
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Tuple containing cancel ID and array of per-item error codes
    fn write_vqt(
        &self,
        server_handles: &[u32],
        values: &[opc_da_bindings::tagOPCITEMVQT],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.len() != values.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and values must have the same length",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.WriteVQT(
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

    /// Refreshes all active items with maximum age constraint.
    ///
    /// # Arguments
    /// * `max_age` - Maximum age constraint in milliseconds
    /// * `transaction_id` - Client-provided transaction identifier
    ///
    /// # Returns
    /// Cancel ID for the refresh operation
    fn refresh_max_age(&self, max_age: u32, transaction_id: u32) -> windows::core::Result<u32> {
        unsafe { self.interface()?.RefreshMaxAge(max_age, transaction_id) }
    }
}
