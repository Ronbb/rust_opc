use crate::client::memory::{LocalPointer, RemoteArray, RemotePointer};
use opc_da_bindings::IOPCItemProperties;
use std::str::FromStr;

pub trait ItemPropertiesTrait {
    fn interface(&self) -> &IOPCItemProperties;

    fn query_available_properties(
        &self,
        item_id: &str,
    ) -> windows::core::Result<(
        RemoteArray<u32>,                 // property IDs
        RemoteArray<windows_core::PWSTR>, // descriptions
        RemoteArray<u16>,                 // datatypes
    )> {
        let item_id = LocalPointer::from_str(item_id)?;

        let mut count = 0;
        let mut property_ids = RemoteArray::new(0);
        let mut descriptions = RemoteArray::new(0);
        let mut datatypes = RemoteArray::new(0);

        unsafe {
            self.interface().QueryAvailableProperties(
                item_id.as_pcwstr(),
                &mut count,
                property_ids.as_mut_ptr(),
                descriptions.as_mut_ptr(),
                datatypes.as_mut_ptr(),
            )?;
        }

        if count > 0 {
            property_ids.set_len(count as _);
            descriptions.set_len(count as _);
            datatypes.set_len(count as _);
        }

        Ok((property_ids, descriptions, datatypes))
    }

    fn get_item_properties(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<windows_core::VARIANT>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if property_ids.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "property_ids is empty",
            ));
        }

        let item_id = LocalPointer::from_str(item_id)?;

        let mut values = RemoteArray::new(property_ids.len());
        let mut errors = RemoteArray::new(property_ids.len());

        unsafe {
            self.interface().GetItemProperties(
                item_id.as_pcwstr(),
                property_ids.len() as u32,
                property_ids.as_ptr(),
                values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((values, errors))
    }

    fn lookup_item_ids(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> windows::core::Result<(
        RemoteArray<windows_core::PWSTR>,
        RemoteArray<windows::core::HRESULT>,
    )> {
        if property_ids.is_empty() {
            return Err(windows_core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "property_ids is empty",
            ));
        }

        let item_id = LocalPointer::from_str(item_id)?;

        let mut new_item_ids = RemoteArray::new(property_ids.len());
        let mut errors = RemoteArray::new(property_ids.len());

        unsafe {
            self.interface().LookupItemIDs(
                item_id.as_pcwstr(),
                property_ids.len() as u32,
                property_ids.as_ptr(),
                new_item_ids.as_mut_ptr(),
                errors.as_mut_ptr(),
            )?;
        }

        Ok((new_item_ids, errors))
    }
}
