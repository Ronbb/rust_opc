use crate::client::memory::Array;

pub trait ItemDeadbandMgtTrait {
    fn item_deadband_mgt(&self) -> &opc_da_bindings::IOPCItemDeadbandMgt;

    fn set_item_deadband(
        &self,
        server_handles: &[u32],
        deadbands: &[f32],
    ) -> windows::core::Result<Array<windows::core::HRESULT>> {
        if server_handles.len() != deadbands.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and deadbands must have the same length",
            ));
        }

        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.item_deadband_mgt().SetItemDeadband(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                deadbands.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn get_item_deadband(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<(Array<f32>, Array<windows::core::HRESULT>)> {
        let mut errors = Array::new(server_handles.len());
        let mut deadbands = Array::new(server_handles.len());

        unsafe {
            self.item_deadband_mgt().GetItemDeadband(
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
    ) -> windows::core::Result<Array<windows::core::HRESULT>> {
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.item_deadband_mgt().ClearItemDeadband(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }
}
