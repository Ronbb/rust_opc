use crate::{
    server::traits::{ItemOptionalVqt, ItemWithMaxAge, ServerTrait},
    utils::{TryToLocal as _, TryToNative as _},
};

use super::{
    enumeration::{ConnectionPointsEnumerator, StringEnumerator},
    utils::{
        PointerReader, PointerWriter, TryReadArray, TryWrite, TryWriteArrayPointer, TryWriteInto,
        TryWriteTo,
    },
};

#[windows::core::implement(
    // implicit implement IUnknown
    opc_da_bindings::IOPCServer,
    opc_da_bindings::IOPCCommon,
    windows::Win32::System::Com::IConnectionPointContainer,
    opc_da_bindings::IOPCItemProperties,
    opc_da_bindings::IOPCBrowse,
    opc_da_bindings::IOPCServerPublicGroups,
    opc_da_bindings::IOPCBrowseServerAddressSpace,
    opc_da_bindings::IOPCItemIO
)]
pub struct Server<T>(pub T)
where
    T: ServerTrait + 'static;

impl<T: ServerTrait + 'static> core::ops::Deref for Server<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCServer_Impl for Server_Impl<T> {
    fn AddGroup(
        &self,
        name: &windows::core::PCWSTR,
        active: windows::Win32::Foundation::BOOL,
        requested_update_rate: u32,
        client_group: u32,
        time_bias: *const i32,
        percent_deadband: *const f32,
        locale_id: u32,
        server_group: *mut u32,
        revised_update_rate: *mut u32,
        reference_interface_id: *const windows::core::GUID,
        unknown: *mut Option<windows::core::IUnknown>,
    ) -> windows::core::Result<()> {
        let info = self.add_group(
            unsafe { name.to_string() }?,
            active.as_bool(),
            requested_update_rate,
            client_group,
            unsafe { time_bias.as_ref() }.copied(),
            unsafe { percent_deadband.as_ref() }.copied(),
            locale_id,
            unsafe { reference_interface_id.as_ref() }.map(|id| id.to_u128()),
        )?;

        PointerWriter::try_write(info.server_group, server_group)?;
        PointerWriter::try_write(info.revised_update_rate, revised_update_rate)?;
        PointerWriter::try_write_into(info.unknown, unknown)?;

        Ok(())
    }

    fn GetErrorString(
        &self,
        error: windows::core::HRESULT,
        locale: u32,
    ) -> windows::core::Result<windows::core::PWSTR> {
        let s = self.get_error_string_locale(error.0, locale)?;
        PointerWriter::try_write_to(&s)
    }

    fn GetGroupByName(
        &self,
        name: &windows::core::PCWSTR,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        self.get_group_by_name(
            unsafe { name.to_string() }?,
            if reference_interface_id.is_null() {
                None
            } else {
                Some(unsafe { *reference_interface_id }.to_u128())
            },
        )
    }

    fn GetStatus(&self) -> windows::core::Result<*mut opc_da_bindings::tagOPCSERVERSTATUS> {
        let status: opc_da_bindings::tagOPCSERVERSTATUS = self.get_status()?.try_into()?;
        PointerWriter::try_write_to(status)
    }

    fn RemoveGroup(
        &self,
        server_group: u32,
        force: windows::Win32::Foundation::BOOL,
    ) -> windows::core::Result<()> {
        self.remove_group(server_group, force.as_bool())
    }

    fn CreateGroupEnumerator(
        &self,
        scope: opc_da_bindings::tagOPCENUMSCOPE,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        self.create_group_enumerator(
            scope.try_to_local()?,
            unsafe { reference_interface_id.as_ref() }.map(|id| id.to_u128()),
        )
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCCommon_Impl for Server_Impl<T> {
    fn SetLocaleID(&self, locale_id: u32) -> windows::core::Result<()> {
        self.set_locale_id(locale_id)
    }

    fn GetLocaleID(&self) -> windows::core::Result<u32> {
        self.get_locale_id()
    }

    fn QueryAvailableLocaleIDs(
        &self,
        count: *mut u32,
        locale_ids: *mut *mut u32,
    ) -> windows::core::Result<()> {
        let available_locale_ids = self.query_available_locale_ids()?;
        PointerWriter::try_write(available_locale_ids.len() as _, count)?;
        PointerWriter::try_write_array_pointer(&available_locale_ids, locale_ids)?;
        Ok(())
    }

    fn GetErrorString(
        &self,
        error: windows::core::HRESULT,
    ) -> windows::core::Result<windows::core::PWSTR> {
        let s = self.get_error_string(error.0).map_err(|e| {
            // Map internal errors to appropriate COM errors
            windows::core::Error::new(windows::Win32::Foundation::E_FAIL, e.to_string())
        })?;
        let mut out = windows::core::PWSTR::null();
        PointerWriter::try_write_into(&s, &mut out).map_err(|e| {
            // Handle allocation failures
            windows::core::Error::new(windows::Win32::Foundation::E_OUTOFMEMORY, e.to_string())
        })?;
        Ok(out)
    }

    fn SetClientName(&self, name: &windows::core::PCWSTR) -> windows::core::Result<()> {
        self.set_client_name(unsafe { name.to_string() }?)
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> windows::Win32::System::Com::IConnectionPointContainer_Impl
    for Server_Impl<T>
{
    fn EnumConnectionPoints(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints> {
        let connection_points = self.enum_connection_points()?;

        Ok(
            windows::core::ComObjectInner::into_object(ConnectionPointsEnumerator::new(
                connection_points,
            ))
            .into_interface(),
        )
    }

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint> {
        self.find_connection_point(reference_interface_id)
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 N/A
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCItemProperties_Impl for Server_Impl<T> {
    fn QueryAvailableProperties(
        &self,
        item_id: &windows::core::PCWSTR,
        count: *mut u32,
        property_ids: *mut *mut u32,
        descriptions: *mut *mut windows::core::PWSTR,
        data_types: *mut *mut u16,
    ) -> windows::core::Result<()> {
        let vec = self.query_available_properties(unsafe { item_id.to_string() }?)?;

        PointerWriter::try_write(vec.len() as _, count)?;

        PointerWriter::try_write_array_pointer(
            &vec.iter().map(|p| p.property_id).collect::<Vec<_>>(),
            property_ids,
        )?;

        PointerWriter::try_write_into(
            &vec.iter()
                .map(|p| p.description.as_str())
                .collect::<Vec<_>>(),
            descriptions,
        )?;

        PointerWriter::try_write_array_pointer(
            &vec.iter().map(|p| p.data_type).collect::<Vec<_>>(),
            data_types,
        )?;

        Ok(())
    }

    fn GetItemProperties(
        &self,
        item_id: &windows::core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        data: *mut *mut windows::Win32::System::Variant::VARIANT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        let property_ids = PointerReader::try_read_array(count, property_ids)?;

        let vec = self.get_item_properties(unsafe { item_id.to_string() }?, property_ids)?;

        PointerWriter::try_write_array_pointer(
            &vec.iter().map(|p| p.error).collect::<Vec<_>>(),
            errors,
        )?;

        PointerWriter::try_write_array_pointer(
            &vec.into_iter().map(|p| p.data.into()).collect::<Vec<_>>(),
            data,
        )?;

        Ok(())
    }

    fn LookupItemIDs(
        &self,
        item_id: &windows::core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        new_item_ids: *mut *mut windows::core::PWSTR,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        let property_ids = PointerReader::try_read_array(count, property_ids)?;

        let vec = self.lookup_item_ids(unsafe { item_id.to_string() }?, property_ids)?;

        PointerWriter::try_write_into(
            &vec.iter()
                .map(|p| p.new_item_id.as_str())
                .collect::<Vec<_>>(),
            new_item_ids,
        )?;

        PointerWriter::try_write_array_pointer(
            &vec.iter().map(|p| p.error).collect::<Vec<_>>(),
            errors,
        )?;

        Ok(())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCBrowse_Impl for Server_Impl<T> {
    fn GetProperties(
        &self,
        item_count: u32,
        item_ids: *const windows::core::PCWSTR,
        return_property_values: windows::Win32::Foundation::BOOL,
        property_count: u32,
        property_ids: *const u32,
        item_properties: *mut *mut opc_da_bindings::tagOPCITEMPROPERTIES,
    ) -> windows::core::Result<()> {
        let item_ids = PointerReader::try_read_array(item_count, item_ids)?;
        let property_ids = PointerReader::try_read_array(property_count, property_ids)?;

        let properties =
            self.get_properties(item_ids, return_property_values.as_bool(), property_ids)?;

        PointerWriter::try_write_array_pointer(
            &properties
                .into_iter()
                .map(|item| match item.try_into() {
                    Ok(item) => item,
                    Err(error) => opc_da_bindings::tagOPCITEMPROPERTIES {
                        hrErrorID: (error as windows::core::Error).code(),
                        ..Default::default()
                    },
                })
                .collect::<Vec<_>>(),
            item_properties,
        )?;

        Ok(())
    }

    fn Browse(
        &self,
        item_id: &windows::core::PCWSTR,
        continuation_point: *mut windows::core::PWSTR,
        max_elements_returned: u32,
        browse_filter: opc_da_bindings::tagOPCBROWSEFILTER,
        element_name_filter: &windows::core::PCWSTR,
        vendor_filter: &windows::core::PCWSTR,
        return_all_properties: windows::Win32::Foundation::BOOL,
        return_property_values: windows::Win32::Foundation::BOOL,
        property_count: u32,
        property_ids: *const u32,
        more_elements: *mut windows::Win32::Foundation::BOOL,
        count: *mut u32,
        browse_elements: *mut *mut opc_da_bindings::tagOPCBROWSEELEMENT,
    ) -> windows::core::Result<()> {
        let item_id = unsafe { item_id.to_string()? };
        let element_name_filter = unsafe { element_name_filter.to_string()? };
        let vendor_filter = unsafe { vendor_filter.to_string()? };
        let property_ids = PointerReader::try_read_array(property_count, property_ids)?;

        let result = self.browse(
            item_id,
            unsafe {
                continuation_point
                    .as_ref()
                    .map(|s| s.to_string())
                    .transpose()?
            },
            max_elements_returned,
            browse_filter.try_into()?,
            element_name_filter,
            vendor_filter,
            return_all_properties.as_bool(),
            return_property_values.as_bool(),
            property_ids,
        )?;

        PointerWriter::try_write(result.elements.len() as _, count)?;

        PointerWriter::try_write_array_pointer(
            &result
                .elements
                .into_iter()
                .map(|element| match element.try_into() {
                    Ok(element) => element,
                    Err(error) => {
                        let mut element = opc_da_bindings::tagOPCBROWSEELEMENT::default();
                        element.ItemProperties.hrErrorID = (error as windows::core::Error).code();
                        element
                    }
                })
                .collect::<Vec<_>>(),
            browse_elements,
        )?;

        PointerWriter::try_write(result.more_elements.into(), more_elements)?;

        match result.continuation_point {
            Some(new_continuation_point) => {
                PointerWriter::try_write_into(&new_continuation_point, continuation_point)?
            }
            None => unsafe {
                *continuation_point = windows::core::PWSTR::null();
            },
        }

        Ok(())
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCServerPublicGroups_Impl for Server_Impl<T> {
    fn GetPublicGroupByName(
        &self,
        name: &windows::core::PCWSTR,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        if reference_interface_id.is_null() {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Null reference_interface_id",
            ));
        }

        self.get_public_group_by_name(
            unsafe { name.to_string() }?,
            unsafe { *reference_interface_id }.to_u128(),
        )
    }

    fn RemovePublicGroup(
        &self,
        server_group: u32,
        force: windows::Win32::Foundation::BOOL,
    ) -> windows::core::Result<()> {
        self.remove_public_group(server_group, force.as_bool())
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCBrowseServerAddressSpace_Impl
    for Server_Impl<T>
{
    fn QueryOrganization(&self) -> windows::core::Result<opc_da_bindings::tagOPCNAMESPACETYPE> {
        self.query_organization().map(Into::into)
    }

    fn ChangeBrowsePosition(
        &self,
        browse_direction: opc_da_bindings::tagOPCBROWSEDIRECTION,
        string: &windows::core::PCWSTR,
    ) -> windows::core::Result<()> {
        self.change_browse_position((browse_direction, unsafe { string.to_string() }?).try_into()?)
    }

    fn BrowseOPCItemIDs(
        &self,
        browse_filter_type: opc_da_bindings::tagOPCBROWSETYPE,
        filter_criteria: &windows::core::PCWSTR,
        variant_data_type_filter: u16,
        access_rights_filter: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString> {
        self.browse_opc_item_ids(
            browse_filter_type.try_into()?,
            unsafe { filter_criteria.to_string() }?,
            variant_data_type_filter,
            access_rights_filter,
        )
    }

    fn GetItemID(
        &self,
        item_data_id: &windows::core::PCWSTR,
    ) -> windows::core::Result<windows::core::PWSTR> {
        let item_id = self.get_item_id(unsafe { item_data_id.to_string() }?)?;
        PointerWriter::try_write_to(&item_id)
    }

    fn BrowseAccessPaths(
        &self,
        item_id: &windows::core::PCWSTR,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumString> {
        let access_paths = self.browse_access_paths(unsafe { item_id.to_string() }?)?;

        Ok(
            windows::core::ComObjectInner::into_object(StringEnumerator::new(access_paths))
                .into_interface(),
        )
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: ServerTrait + 'static> opc_da_bindings::IOPCItemIO_Impl for Server_Impl<T> {
    fn Read(
        &self,
        count: u32,
        item_ids: *const windows::core::PCWSTR,
        max_ages: *const u32,
        values: *mut *mut windows::Win32::System::Variant::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        let item_ids = PointerReader::try_read_array(count, item_ids)?;
        let max_ages = PointerReader::try_read_array(count, max_ages)?;

        let result = self.read(
            item_ids
                .into_iter()
                .zip(max_ages)
                .map(|(item_id, max_age)| ItemWithMaxAge { item_id, max_age })
                .collect(),
        )?;

        PointerWriter::try_write_array_pointer(
            &result
                .iter()
                .map(|vqt| vqt.value.clone().into())
                .collect::<Vec<_>>(),
            values,
        )?;

        PointerWriter::try_write_array_pointer(
            &result.iter().map(|vqt| vqt.quality).collect::<Vec<_>>(),
            qualities,
        )?;

        PointerWriter::try_write_array_pointer(
            &result
                .iter()
                .map(|vqt| vqt.timestamp.try_to_native())
                .collect::<windows::core::Result<Vec<_>>>()?,
            timestamps,
        )?;

        PointerWriter::try_write_array_pointer(
            &result.iter().map(|vqt| vqt.error).collect::<Vec<_>>(),
            errors,
        )?;

        Ok(())
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_ids: *const windows::core::PCWSTR,
        item_vqt: *const opc_da_bindings::tagOPCITEMVQT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        let item_ids = PointerReader::try_read_array(count, item_ids)?;
        let item_vqt = PointerReader::try_read_array(count, item_vqt)?
            .into_iter()
            .try_fold(vec![], |mut acc, item| {
                acc.push(item.try_into()?);
                windows::core::Result::Ok(acc)
            })?;

        let result = self.write_vqt(
            item_ids
                .into_iter()
                .zip(item_vqt)
                .map(|(item_id, optional_vqt)| ItemOptionalVqt {
                    item_id,
                    optional_vqt,
                })
                .collect(),
        )?;

        PointerWriter::try_write_array_pointer(&result, errors)?;

        Ok(())
    }
}
