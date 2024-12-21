use crate::utils::{LocalPointer, RemoteArray};
use opc_da_bindings::{tagOPCITEMVQT, IOPCItemIO};

/// Direct item I/O functionality (OPC DA 3.0).
///
/// Provides methods for direct read/write operations on items without requiring
/// group creation. This trait offers a simplified interface for basic data access.
pub trait ItemIoTrait {
    fn interface(&self) -> windows::core::Result<&IOPCItemIO>;

    /// Reads values directly from items with age constraints.
    ///
    /// # Arguments
    /// * `item_ids` - Array of fully qualified item IDs to read
    /// * `max_age` - Maximum age constraints for each item in milliseconds
    ///
    /// # Returns
    /// Tuple containing:
    /// - Array of item values (VARIANT)
    /// - Array of quality values
    /// - Array of timestamps
    /// - Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    #[allow(clippy::type_complexity)]
    fn read(
        &self,
        item_ids: &[String],
        max_age: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<windows::core::VARIANT>,
        RemoteArray<u16>,
        RemoteArray<windows::Win32::Foundation::FILETIME>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if item_ids.is_empty() || max_age.is_empty() || item_ids.len() != max_age.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid arguments - arrays must be non-empty and have same length",
            ));
        }

        let item_ptrs: LocalPointer<Vec<Vec<u16>>> = LocalPointer::from(item_ids);
        let item_ptrs = item_ptrs.as_pcwstr_array();

        let len = item_ids.len().try_into()?;

        let mut values = RemoteArray::new(len);
        let mut qualities = RemoteArray::new(len);
        let mut timestamps = RemoteArray::new(len);
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Read(
                item_ids.len() as u32,
                item_ptrs.as_ptr(),
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
    /// * `item_ids` - Array of fully qualified item IDs to write
    /// * `item_vqts` - Array of value-quality-timestamp structures
    ///
    /// # Returns
    /// Array of per-item error codes
    ///
    /// # Errors
    /// Returns E_INVALIDARG if arrays are empty or have different lengths
    fn write_vqt(
        &self,
        item_ids: &[String],
        item_vqts: &[tagOPCITEMVQT],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if item_ids.is_empty() || item_vqts.is_empty() || item_ids.len() != item_vqts.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid arguments - arrays must be non-empty and have same length",
            ));
        }

        let len = item_ids.len().try_into()?;

        let item_ptrs = LocalPointer::from(item_ids);
        let item_ptrs = item_ptrs.as_pcwstr_array();

        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.WriteVQT(
                len,
                item_ptrs.as_ptr(),
                item_vqts.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
