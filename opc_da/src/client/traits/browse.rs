use crate::client::memory::{LocalPointer, RemoteArray, RemotePointer};
use opc_da_bindings::{tagOPCBROWSEELEMENT, tagOPCBROWSEFILTER, tagOPCITEMPROPERTIES, IOPCBrowse};
use std::str::FromStr;

pub trait BrowseTrait {
    fn interface(&self) -> windows::core::Result<&IOPCBrowse>;

    fn get_properties(
        &self,
        item_ids: &[String],
        return_property_values: bool,
        property_ids: &[u32],
    ) -> windows::core::Result<RemoteArray<tagOPCITEMPROPERTIES>> {
        if item_ids.is_empty() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "item_ids is empty",
            ));
        }

        let item_ptrs: LocalPointer<Vec<Vec<u16>>> = LocalPointer::from(item_ids);
        let item_ptrs = item_ptrs.as_pcwstr_array();

        let mut results = RemoteArray::new(item_ids.len().try_into()?);

        unsafe {
            self.interface()?.GetProperties(
                item_ids.len() as u32,
                item_ptrs.as_ptr(),
                windows::Win32::Foundation::BOOL::from(return_property_values),
                property_ids,
                results.as_mut_ptr(),
            )?;
        }

        Ok(results)
    }

    #[allow(clippy::too_many_arguments)]
    fn browse(
        &self,
        item_id: &str,
        max_elements: u32,
        browse_filter: tagOPCBROWSEFILTER,
        element_name_filter: &str,
        vendor_filter: &str,
        return_all_properties: bool,
        return_property_values: bool,
        property_ids: &[u32],
    ) -> windows::core::Result<(bool, RemoteArray<tagOPCBROWSEELEMENT>)> {
        let item_id = LocalPointer::from_str(item_id)?;
        let element_name_filter = LocalPointer::from_str(element_name_filter)?;
        let vendor_filter = LocalPointer::from_str(vendor_filter)?;
        let mut continuation_point = RemotePointer::<u16>::new();
        let mut more_elements = false.into();
        let mut count = 0;
        let mut elements = RemoteArray::empty();

        unsafe {
            self.interface()?.Browse(
                item_id.as_pcwstr(),
                continuation_point.as_mut_pwstr_ptr(),
                max_elements,
                browse_filter,
                element_name_filter.as_pcwstr(),
                vendor_filter.as_pcwstr(),
                return_all_properties,
                return_property_values,
                property_ids,
                &mut more_elements,
                &mut count,
                elements.as_mut_ptr(),
            )?;
        }

        if count > 0 {
            elements.set_len(count);
        }

        Ok((more_elements.into(), elements))
    }
}
