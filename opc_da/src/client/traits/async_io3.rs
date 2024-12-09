use crate::client::memory::RemoteArray;

pub trait AsyncIo3Trait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO3>;

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

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(server_handles.len().try_into()?);

        unsafe {
            self.interface()?.ReadMaxAge(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                max_age.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

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

        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(server_handles.len().try_into()?);

        unsafe {
            self.interface()?.WriteVQT(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                values.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

    fn refresh_max_age(&self, max_age: u32, transaction_id: u32) -> windows::core::Result<u32> {
        unsafe { self.interface()?.RefreshMaxAge(max_age, transaction_id) }
    }
}
