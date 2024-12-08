use crate::client::memory::Array;
use windows::core::VARIANT;

pub trait AsyncIo2Trait {
    fn async_io2(&self) -> &opc_da_bindings::IOPCAsyncIO2;

    fn read(
        &self,
        server_handles: &[u32],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, Array<windows::core::HRESULT>)> {
        let mut cancel_id = 0;
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.async_io2().Read(
                server_handles.len() as u32,
                server_handles.as_ptr(),
                transaction_id,
                &mut cancel_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((cancel_id, errors))
    }

    fn write(
        &self,
        server_handles: &[u32],
        values: &[VARIANT],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, Array<windows::core::HRESULT>)> {
        let mut cancel_id = 0;
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.async_io2().Write(
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

    fn refresh2(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        transaction_id: u32,
    ) -> windows::core::Result<u32> {
        unsafe { self.async_io2().Refresh2(source, transaction_id) }
    }

    fn cancel2(&self, cancel_id: u32) -> windows::core::Result<()> {
        unsafe { self.async_io2().Cancel2(cancel_id) }
    }

    fn set_enable(&self, enable: bool) -> windows::core::Result<()> {
        unsafe {
            self.async_io2()
                .SetEnable(windows::Win32::Foundation::BOOL::from(enable))
        }
    }

    fn get_enable(&self) -> windows::core::Result<bool> {
        unsafe { self.async_io2().GetEnable().map(|v| v.as_bool()) }
    }
}
