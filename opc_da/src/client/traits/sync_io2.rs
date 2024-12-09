use crate::client::memory::RemoteArray;
use windows::core::VARIANT;

pub trait SyncIo2Trait {
    fn interface(&self) -> windows_core::Result<&opc_da_bindings::IOPCSyncIO2>;

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

        let mut values = RemoteArray::new(server_handles.len());
        let mut qualities = RemoteArray::new(server_handles.len());
        let mut timestamps = RemoteArray::new(server_handles.len());
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface()?.ReadMaxAge(
                server_handles.len() as u32,
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

        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface()?.WriteVQT(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                values.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
