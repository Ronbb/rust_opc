use crate::client::memory::Array;
use windows::core::VARIANT;

pub trait SyncIoTrait {
    fn sync_io(&self) -> &opc_da_bindings::IOPCSyncIO;

    fn read(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(
        Array<opc_da_bindings::tagOPCITEMSTATE>,
        Array<windows::core::HRESULT>,
    )> {
        let mut item_values = Array::new(server_handles.len());
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.sync_io().Read(
                source,
                server_handles.len() as u32,
                server_handles.as_ptr(),
                item_values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((item_values, errors))
    }

    fn write(
        &self,
        server_handles: &[u32],
        values: &[VARIANT],
    ) -> windows::core::Result<Array<windows::core::HRESULT>> {
        if server_handles.len() != values.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and values must have the same length",
            ));
        }

        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.sync_io().Write(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                values.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
