use opc_classic_types::{ComObject, Result, ValueType};
use opc_classic_utils::{CoTaskMemArrayOut, DropElements, FreePwstrElements, NoCleanup};
use windows_core::PCWSTR;

use crate::IOPCItemProperties;
use crate::abi::{
    error_code_from_abi, from_abi_error, interface_from_object, object_from_interface,
    value_from_abi,
};

use super::group::{count, wide};
use super::{ItemError, PropertyDescription, PropertyValue};

#[derive(Clone)]
pub struct ItemPropertiesClient {
    inner: IOPCItemProperties,
}

impl ItemPropertiesClient {
    pub fn from_object(object: &ComObject) -> Result<Self> {
        interface_from_object(object).map(Self::from_interface)
    }

    pub(crate) fn from_interface(inner: IOPCItemProperties) -> Self {
        Self { inner }
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn available(&self, item_id: &str) -> Result<Vec<PropertyDescription>> {
        let item_id = wide(item_id)?;
        let mut count_value = 0u32;
        let mut ids = CoTaskMemArrayOut::new(0, NoCleanup);
        let mut descriptions = CoTaskMemArrayOut::new(0, FreePwstrElements);
        let mut data_types = CoTaskMemArrayOut::new(0, NoCleanup);
        let call = unsafe {
            self.inner.QueryAvailableProperties(
                PCWSTR(item_id.as_ptr()),
                &mut count_value,
                ids.as_mut_ptr(),
                descriptions.as_mut_ptr().cast(),
                data_types.as_mut_ptr(),
            )
        };
        let len = count_value as usize;
        unsafe {
            ids.set_len(len);
            descriptions.set_len(len);
            data_types.set_len(len);
        }
        let ids = unsafe { ids.into_array() };
        let descriptions = unsafe { descriptions.into_array() };
        let data_types = unsafe { data_types.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let descriptions = descriptions?;
        let data_types = data_types?;
        (0..len)
            .map(|index| {
                Ok(PropertyDescription {
                    id: ids.as_slice()[index],
                    description: pwstr_to_string(descriptions.as_slice()[index])?,
                    data_type: ValueType::from_raw(data_types.as_slice()[index]),
                })
            })
            .collect()
    }

    pub fn values(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> Result<Vec<std::result::Result<PropertyValue, ItemError>>> {
        let item_id = wide(item_id)?;
        let mut values = CoTaskMemArrayOut::new(property_ids.len(), DropElements);
        let mut errors = CoTaskMemArrayOut::new(property_ids.len(), NoCleanup);
        let call = unsafe {
            self.inner.GetItemProperties(
                PCWSTR(item_id.as_ptr()),
                count(property_ids.len())?,
                property_ids.as_ptr(),
                values.as_mut_ptr(),
                errors.as_mut_ptr(),
            )
        };
        let values = unsafe { values.into_array() };
        let errors = unsafe { errors.into_array() };
        call.map_err(from_abi_error)?;
        let values = values?;
        let errors = errors?;
        property_ids
            .iter()
            .zip(values.as_slice().iter().zip(errors.as_slice()))
            .map(|(id, (value, error))| {
                if error.is_err() {
                    Ok(Err(ItemError {
                        code: error_code_from_abi(*error),
                    }))
                } else {
                    Ok(Ok(PropertyValue {
                        id: *id,
                        value: value_from_abi(value)?,
                    }))
                }
            })
            .collect()
    }

    pub fn lookup_item_ids(
        &self,
        item_id: &str,
        property_ids: &[u32],
    ) -> Result<Vec<std::result::Result<String, ItemError>>> {
        let item_id = wide(item_id)?;
        let mut ids = CoTaskMemArrayOut::new(property_ids.len(), FreePwstrElements);
        let mut errors = CoTaskMemArrayOut::new(property_ids.len(), NoCleanup);
        let call = unsafe {
            self.inner.LookupItemIDs(
                PCWSTR(item_id.as_ptr()),
                count(property_ids.len())?,
                property_ids.as_ptr(),
                ids.as_mut_ptr().cast(),
                errors.as_mut_ptr(),
            )
        };
        let ids = unsafe { ids.into_array() };
        let errors = unsafe { errors.into_array() };
        call.map_err(from_abi_error)?;
        let ids = ids?;
        let errors = errors?;
        ids.as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(id, error)| {
                if error.is_err() {
                    Ok(Err(ItemError {
                        code: error_code_from_abi(*error),
                    }))
                } else {
                    pwstr_to_string(*id).map(Ok)
                }
            })
            .collect()
    }
}

fn pwstr_to_string(value: *mut u16) -> Result<String> {
    if value.is_null() {
        Ok(String::new())
    } else {
        // SAFETY: The surrounding owning array guarantees a valid OPC string.
        unsafe { windows_core::PWSTR(value).to_string() }
            .map_err(|_| opc_classic_types::Error::invalid_argument("invalid UTF-16 string"))
    }
}
