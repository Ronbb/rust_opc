use crate::client::memory::RemoteArray;
use windows::core::VARIANT;

pub trait AsyncIo2Trait {
    fn interface(&self) -> &opc_da_bindings::IOPCAsyncIO2;

    fn read(
        &self,
        server_handles: &[u32],
        transaction_id: u32,
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().Read(
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
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        let mut cancel_id = 0;
        let mut errors = RemoteArray::new(server_handles.len());

        unsafe {
            self.interface().Write(
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
        unsafe { self.interface().Refresh2(source, transaction_id) }
    }

    fn cancel2(&self, cancel_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface().Cancel2(cancel_id) }
    }

    fn set_enable(&self, enable: bool) -> windows::core::Result<()> {
        unsafe {
            self.interface()
                .SetEnable(windows::Win32::Foundation::BOOL::from(enable))
        }
    }

    fn get_enable(&self) -> windows::core::Result<bool> {
        unsafe { self.interface().GetEnable().map(|v| v.as_bool()) }
    }
}
