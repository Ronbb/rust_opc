use crate::utils::RemoteArray;
use windows::Win32::System::Variant::VARIANT;

/// Synchronous I/O functionality (OPC DA 3.0).
///
/// Provides enhanced synchronous read/write operations with support for
/// quality, timestamp, and maximum age constraints.
pub trait SyncIo2Trait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO2>;

    /// Reads values with maximum age constraint.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles
    /// * `max_age` - Maximum age constraints for each item in milliseconds
    ///
    /// # Returns
    /// Tuple containing:
    /// - Array of item values
    /// - Array of quality values
    /// - Array of timestamps
    /// - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    #[allow(clippy::type_complexity)]
    fn read_max_age(
        &self,
        server_handles: &[u32],
        max_age: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<VARIANT>,
        RemoteArray<u16>,
        RemoteArray<windows::Win32::Foundation::FILETIME>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if server_handles.len() != max_age.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and max_age must have the same length",
            ));
        }

        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut values = RemoteArray::new(len);
        let mut qualities = RemoteArray::new(len);
        let mut timestamps = RemoteArray::new(len);
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.ReadMaxAge(
                len,
                server_handles.as_ptr(),
                max_age.as_ptr(),
                values.as_mut_ptr(),
                qualities.as_mut_ptr(),
                timestamps.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((values, qualities, timestamps, errors))
    }

    /// Writes values with quality and timestamp information.
    ///
    /// # Arguments
    /// * `server_handles` - Array of server item handles
    /// * `values` - Array of value-quality-timestamp structures
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn write_vqt(
        &self,
        server_handles: &[u32],
        values: &[opc_da_bindings::tagOPCITEMVQT],
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
            self.interface()?.WriteVQT(
                len,
                server_handles.as_ptr(),
                values.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
