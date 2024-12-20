use crate::client::memory::{LocalPointer, RemoteArray, RemotePointer};
use opc_da_bindings::{tagOPCBROWSEELEMENT, tagOPCBROWSEFILTER, tagOPCITEMPROPERTIES, IOPCBrowse};

/// Server address space browsing functionality (OPC DA 3.0).
///
/// Provides methods to browse the hierarchical namespace of an OPC server
/// and retrieve item properties.
pub trait BrowseTrait {
    fn interface(&self) -> windows::core::Result<&IOPCBrowse>;

    /// Gets properties for specified items from the server.
    ///
    /// # Arguments
    /// * `item_ids` - Array of item identifiers to get properties for
    /// * `return_property_values` - If true, return actual property values; if false, only property metadata
    /// * `property_ids` - Specific property IDs to retrieve; empty array means all properties
    ///
    /// # Returns
    /// Array of item properties containing property values and/or metadata
    ///
    /// # Errors
    /// Returns E_INVALIDARG if item_ids is empty
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

    /// Browses a single branch or leaf in the server's address space.
    ///
    /// # Arguments
    /// * `item_id` - Starting point for browsing (empty string for root)
    /// * `max_elements` - Maximum number of elements to return
    /// * `browse_filter` - Filter specifying what types of elements to return
    /// * `element_name_filter` - Filter string for element names (can contain wildcards)
    /// * `vendor_filter` - Vendor-specific filter string
    /// * `return_all_properties` - If true, return all available properties
    /// * `return_property_values` - If true, return property values; if false, only property metadata
    /// * `property_ids` - Specific property IDs to retrieve when return_all_properties is false
    ///
    /// # Returns
    /// Tuple containing:
    /// - Boolean indicating if more elements are available
    /// - Array of browse elements containing names and properties
    #[allow(clippy::too_many_arguments)]
    fn browse<S0, S1, S2, S3>(
        &self,
        item_id: Option<S0>,
        continuation_point: Option<S1>,
        max_elements: u32,
        browse_filter: tagOPCBROWSEFILTER,
        element_name_filter: Option<S2>,
        vendor_filter: Option<S3>,
        return_all_properties: bool,
        return_property_values: bool,
        property_ids: &[u32],
    ) -> windows::core::Result<(bool, Option<String>, RemoteArray<tagOPCBROWSEELEMENT>)>
    where
        S0: AsRef<str>,
        S1: AsRef<str>,
        S2: AsRef<str>,
        S3: AsRef<str>,
    {
        let item_id = LocalPointer::from_option(item_id);
        let element_name_filter = LocalPointer::from_option(element_name_filter);
        let vendor_filter = LocalPointer::from_option(vendor_filter);
        let mut continuation_point = RemotePointer::from_option(match &continuation_point {
            Some(point) => Some(point.as_ref()),
            None => None,
        });
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
            unsafe { elements.set_len(count) };
        }

        Ok((
            more_elements.into(),
            continuation_point.try_into()?,
            elements,
        ))
    }
}
