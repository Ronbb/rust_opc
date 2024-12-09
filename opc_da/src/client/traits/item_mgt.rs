use windows_core::Interface as _;

use crate::client::memory;

pub trait ItemMgtTrait {
    fn interface(&self) -> &opc_da_bindings::IOPCItemMgt;

    fn add_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
    ) -> windows_core::Result<(
        memory::RemoteArray<opc_da_bindings::tagOPCITEMRESULT>,
        memory::RemoteArray<windows_core::HRESULT>,
    )> {
        let len = items.len();
        let mut results = memory::RemoteArray::new(len);
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().AddItems(
                len.try_into()?,
                items.as_ptr(),
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((results, errors))
    }

    fn validate_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
        blob_update: bool,
    ) -> windows_core::Result<(
        memory::RemoteArray<opc_da_bindings::tagOPCITEMRESULT>,
        memory::RemoteArray<windows_core::HRESULT>,
    )> {
        let len = items.len();
        let mut results = memory::RemoteArray::new(len);
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().ValidateItems(
                len.try_into()?,
                items.as_ptr(),
                windows::Win32::Foundation::BOOL::from(blob_update),
                results.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((results, errors))
    }

    fn remove_items(
        &self,
        server_handles: &[u32],
    ) -> windows_core::Result<memory::RemoteArray<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().RemoveItems(
                len.try_into()?,
                server_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn set_active_state(
        &self,
        server_handles: &[u32],
        active: bool,
    ) -> windows_core::Result<memory::RemoteArray<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().SetActiveState(
                len.try_into()?,
                server_handles.as_ptr(),
                windows::Win32::Foundation::BOOL::from(active),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn set_client_handles(
        &self,
        server_handles: &[u32],
        client_handles: &[u32],
    ) -> windows_core::Result<memory::RemoteArray<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().SetClientHandles(
                len.try_into()?,
                server_handles.as_ptr(),
                client_handles.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn set_datatypes(
        &self,
        server_handles: &[u32],
        requested_datatypes: &[u16],
    ) -> windows_core::Result<memory::RemoteArray<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::RemoteArray::new(len);

        unsafe {
            self.interface().SetDatatypes(
                len.try_into()?,
                server_handles.as_ptr(),
                requested_datatypes.as_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok(errors)
    }

    fn create_enumerator(
        &self,
        id: &windows_core::GUID,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumUnknown> {
        let enumerator = unsafe { self.interface().CreateEnumerator(id)? };

        enumerator.cast()
    }
}

// ...existing code...
