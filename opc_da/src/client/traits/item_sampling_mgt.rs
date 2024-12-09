use crate::client::memory::RemoteArray;
use windows::Win32::Foundation::BOOL;

pub trait ItemSamplingMgtTrait {
    fn interface(&self) -> &opc_da_bindings::IOPCItemSamplingMgt;

    fn set_item_sampling_rate(
        &self,
        server_handles: &[u32],
        sampling_rates: &[u32],
    ) -> windows::core::Result<(RemoteArray<u32>, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.len() != sampling_rates.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and sampling_rates must have the same length",
            ));
        }

        let mut revised_rates = RemoteArray::new(server_handles.len());
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().SetItemSamplingRate(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                sampling_rates.as_ptr(),
                revised_rates.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((revised_rates, errors))
    }

    fn get_item_sampling_rate(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<(RemoteArray<u32>, RemoteArray<windows::core::HRESULT>)> {
        let mut sampling_rates = RemoteArray::new(server_handles.len());
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().GetItemSamplingRate(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                sampling_rates.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((sampling_rates, errors))
    }

    fn clear_item_sampling_rate(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().ClearItemSamplingRate(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn set_item_buffer_enable(
        &self,
        server_handles: &[u32],
        enable: &[bool],
    ) -> windows::core::Result<RemoteArray<windows::core::HRESULT>> {
        if server_handles.len() != enable.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and enable must have the same length",
            ));
        }

        let mut errors = RemoteArray::new(server_handles.len());
        let enable_bool: Vec<BOOL> = enable.iter().map(|&v| BOOL::from(v)).collect();

        unsafe {
            self.interface().SetItemBufferEnable(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                enable_bool.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn get_item_buffer_enable(
        &self,
        server_handles: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<windows::Win32::Foundation::BOOL>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        let mut enable = RemoteArray::new(server_handles.len());
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().GetItemBufferEnable(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                enable.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((enable, errors))
    }
}
