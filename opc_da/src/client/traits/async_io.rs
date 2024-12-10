use crate::client::memory::RemoteArray;
use windows::core::VARIANT;

pub trait AsyncIoTrait {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO>;

    fn read(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
        server_handles: &[u32],
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut transaction_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Read(
                connection,
                source,
                len,
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
    ) -> windows::core::Result<(u32, RemoteArray<windows::core::HRESULT>)> {
        if server_handles.len() != values.len() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles and values must have the same length",
            ));
        }

        if server_handles.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "server_handles cannot be empty",
            ));
        }

        let len = server_handles.len().try_into()?;

        let mut transaction_id = 0;
        let mut errors = RemoteArray::new(len);

        unsafe {
            self.interface()?.Write(
                connection,
                len,
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
        unsafe { self.interface()?.Refresh(connection, source) }
    }

    fn cancel(&self, transaction_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.Cancel(transaction_id) }
    }
}
