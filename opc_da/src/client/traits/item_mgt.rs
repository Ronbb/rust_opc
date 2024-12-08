use windows_core::Interface as _;

use crate::client::memory;

pub trait ItemMgtTrait {
    fn item_mgt(&self) -> &opc_da_bindings::IOPCItemMgt;

    fn add_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
    ) -> windows_core::Result<(
        memory::Array<opc_da_bindings::tagOPCITEMRESULT>,
        memory::Array<windows_core::HRESULT>,
    )> {
        let len = items.len();
        let mut results = memory::Array::new(len);
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().AddItems(
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
        memory::Array<opc_da_bindings::tagOPCITEMRESULT>,
        memory::Array<windows_core::HRESULT>,
    )> {
        let len = items.len();
        let mut results = memory::Array::new(len);
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().ValidateItems(
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
    ) -> windows_core::Result<memory::Array<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().RemoveItems(
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
    ) -> windows_core::Result<memory::Array<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().SetActiveState(
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
    ) -> windows_core::Result<memory::Array<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().SetClientHandles(
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
    ) -> windows_core::Result<memory::Array<windows_core::HRESULT>> {
        let len = server_handles.len();
        let mut errors = memory::Array::new(len);

        unsafe {
            self.item_mgt().SetDatatypes(
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
        let enumerator = unsafe { self.item_mgt().CreateEnumerator(id)? };

        enumerator.cast()
    }
}

// ...existing code...
