use crate::client::memory::RemoteArray;

pub trait ItemDeadbandMgtTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemDeadbandMgt>;

    fn set_item_deadband(
        &self,
        server_handles: &[u32],
        dead_bands: &[f32],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != dead_bands.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and deadbands must have the same length",
            ));
        }

        let mut errors = RemoteArray::new(server_handles.len().try_into()?);

        unsafe {
            self.interface()?.SetItemDeadband(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                dead_bands.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn get_item_deadband(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<(RemoteArray<f32>, RemoteArray<windows::core::HRESULT>)> {
        let mut errors = RemoteArray::new(server_handles.len().try_into()?);
        let mut deadbands = RemoteArray::new(server_handles.len().try_into()?);

        unsafe {
            self.interface()?.GetItemDeadband(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                deadbands.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((deadbands, errors))
    }

    fn clear_item_deadband(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        let mut errors = RemoteArray::new(server_handles.len().try_into()?);

        unsafe {
            self.interface()?.ClearItemDeadband(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
