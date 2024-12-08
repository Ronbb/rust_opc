use crate::client::memory::Array;
use windows::core::VARIANT;

pub trait AsyncIoTrait {
    fn async_io(&self) -> &opc_da_bindings::IOPCAsyncIO;

    fn read(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(u32, Array<windows::core::HRESULT>)> {
        let mut transaction_id = 0;
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.async_io().Read(
                connection,
                source,
                server_handles.len() as u32,
                server_handles.as_ptr(),
                &mut transaction_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((transaction_id, errors))
    }

    fn write(
        &self,
        connection: u32,
        server_handles: &[u32],
        values: &[VARIANT],
    ) -> windows::core::Result<(u32, Array<windows::core::HRESULT>)> {
        let mut transaction_id = 0;
        let mut errors = Array::new(server_handles.len());

        unsafe {
            self.async_io().Write(
                connection,
                server_handles.len() as u32,
                server_handles.as_ptr(),
                values.as_ptr(),
                &mut transaction_id,
                errors.as_mut_ptr(),
            )?;
        }

        Ok((transaction_id, errors))
    }

    fn refresh(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
    ) -> windows::core::Result<u32> {
        unsafe { self.async_io().Refresh(connection, source) }
    }

    fn cancel(&self, transaction_id: u32) -> windows::core::Result<()> {
        unsafe { self.async_io().Cancel(transaction_id) }
    }
}
