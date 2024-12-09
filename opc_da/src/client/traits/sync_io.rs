use crate::client::memory::RemoteArray;
use windows::core::VARIANT;

pub trait SyncIoTrait {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCSyncIO>;

    fn read(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<opc_da_bindings::tagOPCITEMSTATE>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        let mut item_values = RemoteArray::new(server_handles.len());
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface()?.Read(
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
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != values.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and values must have the same length",
            ));
        }

        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface()?.Write(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                values.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
