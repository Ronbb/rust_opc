use crate::client::memory::{LocalPointer, RemoteArray};
use opc_da_bindings::{tagOPCITEMVQT, IOPCItemIO};

pub trait ItemIoTrait {
    fn interface(&self) -> &IOPCItemIO;

    #[allow(clippy::type_complexity)]
    fn read(
        &self,
        item_ids: &[String],
        max_age: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<windows_core::VARIANT>,
        RemoteArray<u16>,
        RemoteArray<windows::Win32::Foundation::FILETIME>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if item_ids.is_empty() || max_age.is_empty() || item_ids.len() != max_age.len() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid arguments - arrays must be non-empty and have same length",
            ));
        }

        let item_ptrs: LocalPointer<Vec<Vec<u16>>> = LocalPointer::from(item_ids);
        let item_ptrs = item_ptrs.as_pcwstr_array();

        let mut values = RemoteArray::new(item_ids.len());
        let mut qualities = RemoteArray::new(item_ids.len());
        let mut timestamps = RemoteArray::new(item_ids.len());
        let mut errors = RemoteArray::new(item_ids.len());

        unsafe {
            self.interface().Read(
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

    fn write_vqt(
        &self,
        item_ids: &[String],
        item_vqts: &[tagOPCITEMVQT],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if item_ids.is_empty() || item_vqts.is_empty() || item_ids.len() != item_vqts.len() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid arguments - arrays must be non-empty and have same length",
            ));
        }

        let item_ptrs = LocalPointer::from(item_ids);
        let item_ptrs = item_ptrs.as_pcwstr_array();

        let mut errors = RemoteArray::new(item_ids.len());

        unsafe {
            self.interface().WriteVQT(
                item_ids.len() as u32,
                item_ptrs.as_ptr(),
                item_vqts.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
