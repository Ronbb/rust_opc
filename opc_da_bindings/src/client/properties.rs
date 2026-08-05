use opc_classic_utils::{CoTaskMemOut, DropElements, FreePwstrElements, NoCleanup};
use windows::Win32::System::Variant::VARIANT;
use windows_core::{HRESULT, Result};

use crate::IOPCItemProperties;

use super::group::{count, wide};
use super::{ItemError, PropertyDescription, PropertyValue};

#[derive(Clone)]
pub struct ItemPropertiesClient {
    inner: IOPCItemProperties,
}

impl ItemPropertiesClient {
    pub fn new(inner: IOPCItemProperties) -> Self {
        Self { inner }
    }

    pub fn available(&self, item_id: &str) -> Result<Vec<PropertyDescription>> {
        let item_id = wide(item_id)?;
        let mut count_value = 0u32;
        let mut ids = CoTaskMemOut::<u32>::new();
        let mut descriptions = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut data_types = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            self.inner.QueryAvailableProperties(
                item_id.as_pcwstr(),
                &mut count_value,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr(),
                data_types.as_mut_ptr(),
            )
        };
        let len = count_value as usize;
        let ids = unsafe { ids.into_array(len, NoCleanup) }?;
        let descriptions = unsafe { descriptions.into_array(len, FreePwstrElements) }?;
        let data_types = unsafe { data_types.into_array(len, NoCleanup) }?;
        call?;
        Ok((0..len)
            .map(|index| PropertyDescription {
                id: ids.as_slice()[index],
                description: pwstr_to_string(descriptions.as_slice()[index]),
                data_type: data_types.as_slice()[index],
            })
            .collect())
    }

    pub fn values(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> Result<Vec<std::result::Result<PropertyValue, ItemError>>> {
        let item_id = wide(item_id)?;
        let mut values = CoTaskMemOut::<VARIANT>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            self.inner.GetItemProperties(
                item_id.as_pcwstr(),
                count(property_ids.len())?,
                property_ids.as_ptr(),
                values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let values = unsafe { values.into_array(property_ids.len(), DropElements) }?;
        let errors = unsafe { errors.into_array(property_ids.len(), NoCleanup) }?;
        call?;
        Ok(property_ids
            .iter()
            .zip(values.as_slice().iter().zip(errors.as_slice()))
            .map(|(id, (value, error))| {
                if error.is_err() {
                    Err(ItemError { code: *error })
                } else {
                    Ok(PropertyValue {
                        id: *id,
                        value: value.clone(),
                    })
                }
            })
            .collect())
    }

    pub fn lookup_item_ids(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> Result<Vec<std::result::Result<String, ItemError>>> {
        let item_id = wide(item_id)?;
        let mut ids = CoTaskMemOut::<windows_core::PWSTR>::new();
        let mut errors = CoTaskMemOut::<HRESULT>::new();
        let call = unsafe {
            self.inner.LookupItemIDs(
                item_id.as_pcwstr(),
                count(property_ids.len())?,
                property_ids.as_ptr(),
                ids.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let ids = unsafe { ids.into_array(property_ids.len(), FreePwstrElements) }?;
        let errors = unsafe { errors.into_array(property_ids.len(), NoCleanup) }?;
        call?;
        Ok(ids
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(id, error)| {
                if error.is_err() {
                    Err(ItemError { code: *error })
                } else {
                    Ok(pwstr_to_string(*id))
                }
            })
            .collect())
    }
}

fn pwstr_to_string(value: windows_core::PWSTR) -> String {
    if value.is_null() {
        String::new()
    } else {
        // SAFETY: The surrounding owning array guarantees a valid OPC string.
        unsafe { value.to_string() }.unwrap_or_default()
    }
}
