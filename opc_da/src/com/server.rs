use std::{mem::ManuallyDrop, ops::Deref};

use windows::Win32::{
    Foundation::{E_INVALIDARG, E_POINTER},
    System::Com::{
        IConnectionPoint, IConnectionPointContainer, IConnectionPointContainer_Impl,
        IEnumConnectionPoints, IEnumString, IEnumUnknown,
    },
};
use windows_core::{implement, ComObjectInner, IUnknownImpl, Interface, PWSTR};

use crate::traits::{
    BrowseDirection, BrowseElement, BrowseFilter, BrowseType, ItemOptionalVqt, ItemProperties,
    ItemProperty, ItemWithMaxAge, NamespaceType, OptionalVqt, ServerTrait,
};

use super::{
    bindings::{self, tagOPCITEMPROPERTIES},
    enumeration::{ConnectionPointsEnumerator, StringEnumerator, UnknownEnumerator},
    group::Group,
    utils::{
        PointerReader, PointerWriter, ReadArray, TryReadArray, Write, WriteArray,
        WriteArrayPointer, WriteInto, WriteTo,
    },
};

#[implement(
    // implicit implement IUnknown
    bindings::IOPCServer,
    bindings::IOPCCommon,
    IConnectionPointContainer,
    bindings::IOPCItemProperties,
    bindings::IOPCBrowse,
    bindings::IOPCServerPublicGroups,
    bindings::IOPCBrowseServerAddressSpace,
    bindings::IOPCItemIO
)]
pub struct Server<T>(pub T)
where
    T: ServerTrait + 'static;

impl<T: ServerTrait + 'static> Deref for Server<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> bindings::IOPCServer_Impl for Server_Impl<T> {
    fn AddGroup(
        &self,
        name: &windows_core::PCWSTR,
        active: windows::Win32::Foundation::BOOL,
        requested_update_rate: u32,
        client_group: u32,
        time_bias: *const i32,
        percent_deadband: *const f32,
        locale_id: u32,
        server_group: *mut u32,
        revised_update_rate: *mut u32,
        reference_interface_id: *const windows_core::GUID,
        unknown: *mut Option<windows_core::IUnknown>,
    ) -> windows_core::Result<()> {
        if server_group.is_null() || revised_update_rate.is_null() || unknown.is_null() {
            return Err(E_POINTER.into());
        }

        unsafe {
            if let Some(percent_deadband) = percent_deadband.as_ref() {
                if *percent_deadband < 0.0 || *percent_deadband > 100.0 {
                    return Err(E_INVALIDARG.into());
                }
            }
        }

        let mut groups = self.group_name_map.blocking_write();
        let group;
        let server_group_handle = self
            .next_server_group_handle
            .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        let update_rate = requested_update_rate;
        let group_name = unsafe { name.to_string() }?;
        unsafe {
            group = Group::new(
                self.core.clone(),
                group_name.clone(),
                active.as_bool(),
                requested_update_rate,
                client_group,
                time_bias.as_ref().cloned(),
                percent_deadband.as_ref().cloned(),
                locale_id,
                server_group_handle,
            );

            *server_group = server_group_handle;
            *revised_update_rate = update_rate;
        }

        if groups.contains_key(&group_name) {
            return Err(E_INVALIDARG.into());
        }

        let group = group.into_object();

        let result = unsafe {
            group.QueryInterface(
                reference_interface_id,
                unknown as *mut *mut ::core::ffi::c_void,
            )
        };

        if result.is_err() {
            return Err(result.into());
        }

        groups.insert(group_name, group.clone());

        self.group_server_handle_map
            .blocking_write()
            .insert(server_group_handle, group);

        Ok(())
    }

    fn GetErrorString(
        &self,
        error: windows_core::HRESULT,
        locale: u32,
    ) -> windows_core::Result<windows_core::PWSTR> {
        let s = self.get_error_string_locale(error.0, locale)?;
        PointerWriter::write_to(&s)
    }

    fn GetGroupByName(
        &self,
        name: &windows_core::PCWSTR,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
        let name = unsafe { name.to_string().unwrap() };
        if let Some(group) = self.group_name_map.blocking_read().get(&name) {
            let mut unknown = None;
            let result = unsafe {
                group.QueryInterface(
                    reference_interface_id,
                    &mut unknown as *mut _ as *mut *mut ::core::ffi::c_void,
                )
            };

            if result.is_err() {
                return Err(result.into());
            }

            return Ok(unknown.unwrap());
        }

        Err(E_INVALIDARG.into())
    }

    fn GetStatus(&self) -> windows_core::Result<*mut bindings::tagOPCSERVERSTATUS> {
        Ok(&mut bindings::tagOPCSERVERSTATUS::default())
    }

    fn RemoveGroup(
        &self,
        server_group: u32,
        _force: windows::Win32::Foundation::BOOL,
    ) -> windows_core::Result<()> {
        let group = self
            .group_server_handle_map
            .blocking_write()
            .remove(&server_group);
        if let Some(group) = group {
            let mut groups = self.group_name_map.blocking_write();
            groups.remove(&group.state.blocking_read().name);
            Ok(())
        } else {
            Err(E_INVALIDARG.into())
        }
    }

    fn CreateGroupEnumerator(
        &self,
        _scope: bindings::tagOPCENUMSCOPE,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
        match unsafe { reference_interface_id.as_ref() } {
            Some(&IEnumString::IID) => {
                let groups = self.group_name_map.blocking_read();
                let strings = groups.keys().cloned().collect();
                Ok(StringEnumerator::new(strings).into_object().cast().unwrap())
            }
            Some(&IEnumUnknown::IID) => {
                let groups = self.group_name_map.blocking_read();
                let items = groups.values().map(|group| group.to_interface()).collect();
                Ok(UnknownEnumerator::new(items).into_object().cast().unwrap())
            }
            _ => Err(E_INVALIDARG.into()),
        }
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> bindings::IOPCCommon_Impl for Server_Impl<T> {
    fn SetLocaleID(&self, locale_id: u32) -> windows_core::Result<()> {
        self.set_locale_id(locale_id)
    }

    fn GetLocaleID(&self) -> windows_core::Result<u32> {
        self.get_locale_id()
    }

    fn QueryAvailableLocaleIDs(
        &self,
        count: *mut u32,
        locale_ids: *mut *mut u32,
    ) -> windows_core::Result<()> {
        let available_locale_ids = self.query_available_locale_ids()?;
        PointerWriter::write(available_locale_ids.len() as _, count);
        PointerWriter::write_array_pointer(&available_locale_ids, locale_ids)?;
        Ok(())
    }

    fn GetErrorString(
        &self,
        error: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR> {
        let s = self.get_error_string(error.0)?;
        let mut out = PWSTR::null();
        PointerWriter::write_into(&s, &mut out)?;
        Ok(out)
    }

    fn SetClientName(&self, name: &windows_core::PCWSTR) -> windows_core::Result<()> {
        self.set_client_name(unsafe { name.to_string() }?)
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: ServerTrait + 'static> IConnectionPointContainer_Impl for Server_Impl<T> {
    fn EnumConnectionPoints(&self) -> windows_core::Result<IEnumConnectionPoints> {
        let connection_points = self.enum_connection_points()?;

        Ok(ConnectionPointsEnumerator::new(connection_points)
            .into_object()
            .into_interface())
    }

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<IConnectionPoint> {
        self.find_connection_point(reference_interface_id)
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 N/A
impl<T: ServerTrait + 'static> bindings::IOPCItemProperties_Impl for Server_Impl<T> {
    fn QueryAvailableProperties(
        &self,
        item_id: &windows_core::PCWSTR,
        count: *mut u32,
        property_ids: *mut *mut u32,
        descriptions: *mut *mut windows_core::PWSTR,
        data_types: *mut *mut u16,
    ) -> windows_core::Result<()> {
        let vec = self.query_available_properties(unsafe { item_id.to_string() }?)?;

        PointerWriter::write(vec.len() as _, count);

        PointerWriter::write_array_pointer(
            &vec.iter().map(|p| p.property_id).collect::<Vec<_>>(),
            property_ids,
        )?;

        PointerWriter::write_into(
            &vec.iter()
                .map(|p| p.description.as_str())
                .collect::<Vec<_>>(),
            descriptions,
        )?;

        PointerWriter::write_array_pointer(
            &vec.iter().map(|p| p.data_type).collect::<Vec<_>>(),
            data_types,
        )?;

        Ok(())
    }

    fn GetItemProperties(
        &self,
        item_id: &windows_core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        data: *mut *mut windows_core::VARIANT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let property_ids = PointerReader::read_array(count, property_ids);

        let vec = self.get_item_properties(unsafe { item_id.to_string() }?, property_ids)?;

        PointerWriter::write_array_pointer(
            &vec.iter().map(|p| p.error).collect::<Vec<_>>(),
            errors,
        )?;

        PointerWriter::write_array_pointer(
            &vec.into_iter().map(|p| p.data.into()).collect::<Vec<_>>(),
            data,
        )?;

        Ok(())
    }

    fn LookupItemIDs(
        &self,
        item_id: &windows_core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        new_item_ids: *mut *mut windows_core::PWSTR,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let property_ids = PointerReader::read_array(count, property_ids);

        let vec = self.lookup_item_ids(unsafe { item_id.to_string() }?, property_ids)?;

        PointerWriter::write_into(
            &vec.iter()
                .map(|p| p.new_item_id.as_str())
                .collect::<Vec<_>>(),
            new_item_ids,
        )?;

        PointerWriter::write_array_pointer(
            &vec.iter().map(|p| p.error).collect::<Vec<_>>(),
            errors,
        )?;

        Ok(())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: ServerTrait + 'static> bindings::IOPCBrowse_Impl for Server_Impl<T> {
    fn GetProperties(
        &self,
        item_count: u32,
        item_ids: *const windows_core::PCWSTR,
        return_property_values: windows::Win32::Foundation::BOOL,
        property_count: u32,
        property_ids: *const u32,
        item_properties: *mut *mut bindings::tagOPCITEMPROPERTIES,
    ) -> windows_core::Result<()> {
        let item_ids = PointerReader::try_read_array(item_count, item_ids)?;
        let property_ids = PointerReader::read_array(property_count, property_ids);

        let properties =
            self.get_properties(item_ids, return_property_values.as_bool(), property_ids)?;

        PointerWriter::write_array_pointer(
            &properties
                .into_iter()
                .map(|item| match item.try_into() {
                    Ok(item) => item,
                    Err(error) => {
                        let mut item = tagOPCITEMPROPERTIES::default();
                        item.hrErrorID = (error as windows_core::Error).code();
                        item
                    }
                })
                .collect::<Vec<_>>(),
            item_properties,
        )?;

        Ok(())
    }

    fn Browse(
        &self,
        item_id: &windows_core::PCWSTR,
        continuation_point: *mut windows_core::PWSTR,
        max_elements_returned: u32,
        browse_filter: bindings::tagOPCBROWSEFILTER,
        element_name_filter: &windows_core::PCWSTR,
        vendor_filter: &windows_core::PCWSTR,
        return_all_properties: windows::Win32::Foundation::BOOL,
        return_property_values: windows::Win32::Foundation::BOOL,
        property_count: u32,
        property_ids: *const u32,
        more_elements: *mut windows::Win32::Foundation::BOOL,
        count: *mut u32,
        browse_elements: *mut *mut bindings::tagOPCBROWSEELEMENT,
    ) -> windows_core::Result<()> {
        let item_id = unsafe { item_id.to_string()? };
        let element_name_filter = unsafe { element_name_filter.to_string()? };
        let vendor_filter = unsafe { vendor_filter.to_string()? };
        let property_ids = PointerReader::read_array(property_count, property_ids);

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

        PointerWriter::write(result.elements.len() as _, count);

        PointerWriter::write_array_pointer(
            &result
                .elements
                .into_iter()
                .map(|element| match element.try_into() {
                    Ok(element) => element,
                    Err(error) => {
                        let mut element = bindings::tagOPCBROWSEELEMENT::default();
                        element.ItemProperties.hrErrorID = (error as windows_core::Error).code();
                        element
                    }
                })
                .collect::<Vec<_>>(),
            browse_elements,
        )?;

        PointerWriter::write(result.more_elements.into(), more_elements)?;

        match result.continuation_point {
            Some(new_continuation_point) => {
                PointerWriter::write_into(&new_continuation_point, continuation_point)?
            }
            None => unsafe {
                *continuation_point = PWSTR::null();
            },
        }

        Ok(())
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: ServerTrait + 'static> bindings::IOPCServerPublicGroups_Impl for Server_Impl<T> {
    fn GetPublicGroupByName(
        &self,
        name: &windows_core::PCWSTR,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
        self.get_public_group_by_name(
            unsafe { name.to_string() }?,
            unsafe { *reference_interface_id }.to_u128(),
        )
    }

    fn RemovePublicGroup(
        &self,
        server_group: u32,
        force: windows::Win32::Foundation::BOOL,
    ) -> windows_core::Result<()> {
        self.remove_public_group(server_group, force.as_bool())
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: ServerTrait + 'static> bindings::IOPCBrowseServerAddressSpace_Impl for Server_Impl<T> {
    fn QueryOrganization(&self) -> windows_core::Result<bindings::tagOPCNAMESPACETYPE> {
        self.query_organization().map(Into::into)
    }

    fn ChangeBrowsePosition(
        &self,
        browse_direction: bindings::tagOPCBROWSEDIRECTION,
        string: &windows_core::PCWSTR,
    ) -> windows_core::Result<()> {
        self.change_browse_position((browse_direction, unsafe { string.to_string() }?).try_into()?)
    }

    fn BrowseOPCItemIDs(
        &self,
        browse_filter_type: bindings::tagOPCBROWSETYPE,
        filter_criteria: &windows_core::PCWSTR,
        variant_data_type_filter: u16,
        access_rights_filter: u32,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumString> {
        self.browse_opc_item_ids(
            browse_filter_type.try_into()?,
            unsafe { filter_criteria.to_string() }?,
            variant_data_type_filter,
            access_rights_filter,
        )
    }

    fn GetItemID(
        &self,
        item_data_id: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::PWSTR> {
        let item_id = self.get_item_id(unsafe { item_data_id.to_string() }?)?;
        PointerWriter::write_to(&item_id)
    }

    fn BrowseAccessPaths(
        &self,
        item_id: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumString> {
        let access_paths = self.browse_access_paths(unsafe { item_id.to_string() }?)?;

        Ok(StringEnumerator::new(access_paths)
            .into_object()
            .into_interface())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: ServerTrait + 'static> bindings::IOPCItemIO_Impl for Server_Impl<T> {
    fn Read(
        &self,
        count: u32,
        item_ids: *const windows_core::PCWSTR,
        max_ages: *const u32,
        values: *mut *mut windows_core::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let item_ids = PointerReader::try_read_array(count, item_ids)?;
        let max_ages = PointerReader::read_array(count, max_ages);

        let result = self.read(
            item_ids
                .into_iter()
                .zip(max_ages)
                .map(|(item_id, max_age)| ItemWithMaxAge { item_id, max_age })
                .collect(),
        )?;

        PointerWriter::write_array_pointer(
            &result
                .iter()
                .map(|vqt| vqt.value.clone().into())
                .collect::<Vec<_>>(),
            values,
        )?;

        PointerWriter::write_array_pointer(
            &result.iter().map(|vqt| vqt.quality).collect::<Vec<_>>(),
            qualities,
        )?;

        PointerWriter::write_array_pointer(
            &result
                .iter()
                .map(|vqt| vqt.timestamp.clone().into())
                .collect::<Vec<_>>(),
            timestamps,
        )?;

        PointerWriter::write_array_pointer(
            &result.iter().map(|vqt| vqt.error).collect::<Vec<_>>(),
            errors,
        )?;

        Ok(())
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_ids: *const windows_core::PCWSTR,
        item_vqt: *const bindings::tagOPCITEMVQT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let item_ids = PointerReader::try_read_array(count, item_ids)?;
        let item_vqt = PointerReader::read_array(count, item_vqt)
            .into_iter()
            .try_fold(vec![], |mut acc, item| {
                acc.push(item.try_into()?);
                windows_core::Result::Ok(acc)
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

        PointerWriter::write_array_pointer(&result, errors)?;

        Ok(())
    }
}

impl TryFrom<ItemProperties> for opc_da_bindings::tagOPCITEMPROPERTIES {
    type Error = windows_core::Error;

    fn try_from(value: ItemProperties) -> Result<Self, Self::Error> {
        let result = opc_da_bindings::tagOPCITEMPROPERTIES {
            hrErrorID: value.error_id,
            dwNumProperties: value.item_properties.len() as u32,
            pItemProperties: std::ptr::null_mut(),
            dwReserved: 0,
        };

        PointerWriter::write_array(
            &value
                .item_properties
                .into_iter()
                .map(|item_property| match item_property.try_into() {
                    Ok(item_property) => item_property,
                    Err(error) => {
                        let mut result = opc_da_bindings::tagOPCITEMPROPERTY::default();
                        result.hrErrorID = (error as windows_core::Error).code();
                        result
                    }
                })
                .collect::<Vec<_>>(),
            result.pItemProperties,
        )?;

        Ok(result)
    }
}

impl TryFrom<ItemProperty> for opc_da_bindings::tagOPCITEMPROPERTY {
    type Error = windows_core::Error;

    fn try_from(value: ItemProperty) -> Result<Self, Self::Error> {
        Ok(opc_da_bindings::tagOPCITEMPROPERTY {
            vtDataType: value.data_type,
            wReserved: 0,
            dwPropertyID: value.property_id,
            szItemID: PointerWriter::write_to(&value.item_id)?,
            szDescription: PointerWriter::write_to(&value.description)?,
            vValue: ManuallyDrop::new(value.value.into()),
            hrErrorID: value.error_id,
            dwReserved: 0,
        })
    }
}

impl From<BrowseFilter> for opc_da_bindings::tagOPCBROWSEFILTER {
    fn from(value: BrowseFilter) -> Self {
        match value {
            BrowseFilter::All => opc_da_bindings::OPC_BROWSE_FILTER_ALL,
            BrowseFilter::Branches => opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES,
            BrowseFilter::Items => opc_da_bindings::OPC_BROWSE_FILTER_ITEMS,
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCBROWSEFILTER> for BrowseFilter {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCBROWSEFILTER) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_BROWSE_FILTER_ALL => Ok(BrowseFilter::All),
            opc_da_bindings::OPC_BROWSE_FILTER_BRANCHES => Ok(BrowseFilter::Branches),
            opc_da_bindings::OPC_BROWSE_FILTER_ITEMS => Ok(BrowseFilter::Items),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseFilter",
            )),
        }
    }
}

impl TryFrom<BrowseElement> for opc_da_bindings::tagOPCBROWSEELEMENT {
    type Error = windows_core::Error;

    fn try_from(value: BrowseElement) -> Result<Self, Self::Error> {
        Ok(opc_da_bindings::tagOPCBROWSEELEMENT {
            szName: PointerWriter::write_to(&value.name)?,
            szItemID: PointerWriter::write_to(&value.item_id)?,
            dwFlagValue: value.flag_value,
            dwReserved: 0,
            ItemProperties: value.item_properties.try_into()?,
        })
    }
}

impl From<NamespaceType> for opc_da_bindings::tagOPCNAMESPACETYPE {
    fn from(value: NamespaceType) -> Self {
        match value {
            NamespaceType::Flat => opc_da_bindings::OPC_NS_FLAT,
            NamespaceType::Hierarchical => opc_da_bindings::OPC_NS_HIERARCHIAL,
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCNAMESPACETYPE> for NamespaceType {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCNAMESPACETYPE) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_NS_FLAT => Ok(NamespaceType::Flat),
            opc_da_bindings::OPC_NS_HIERARCHIAL => Ok(NamespaceType::Hierarchical),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid NamespaceType",
            )),
        }
    }
}

impl TryFrom<(opc_da_bindings::tagOPCBROWSEDIRECTION, String)> for BrowseDirection {
    type Error = windows_core::Error;

    fn try_from(
        value: (opc_da_bindings::tagOPCBROWSEDIRECTION, String),
    ) -> Result<Self, Self::Error> {
        match value {
            (opc_da_bindings::OPC_BROWSE_UP, _) => Ok(BrowseDirection::Up),
            (opc_da_bindings::OPC_BROWSE_DOWN, _) => Ok(BrowseDirection::Down),
            (opc_da_bindings::OPC_BROWSE_TO, name) => Ok(BrowseDirection::To(name)),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseDirection",
            )),
        }
    }
}

impl From<BrowseDirection> for (opc_da_bindings::tagOPCBROWSEDIRECTION, String) {
    fn from(value: BrowseDirection) -> Self {
        match value {
            BrowseDirection::Up => (opc_da_bindings::OPC_BROWSE_UP, String::new()),
            BrowseDirection::Down => (opc_da_bindings::OPC_BROWSE_DOWN, String::new()),
            BrowseDirection::To(name) => (opc_da_bindings::OPC_BROWSE_TO, name),
        }
    }
}

impl TryFrom<opc_da_bindings::tagOPCBROWSETYPE> for BrowseType {
    type Error = windows_core::Error;

    fn try_from(value: opc_da_bindings::tagOPCBROWSETYPE) -> Result<Self, Self::Error> {
        match value {
            opc_da_bindings::OPC_BRANCH => Ok(BrowseType::Branch),
            opc_da_bindings::OPC_LEAF => Ok(BrowseType::Leaf),
            opc_da_bindings::OPC_FLAT => Ok(BrowseType::Flat),
            _ => Err(windows_core::Error::new(
                E_INVALIDARG,
                "Invalid BrowseFilter",
            )),
        }
    }
}

impl From<BrowseType> for opc_da_bindings::tagOPCBROWSETYPE {
    fn from(value: BrowseType) -> Self {
        match value {
            BrowseType::Branch => opc_da_bindings::OPC_BRANCH,
            BrowseType::Leaf => opc_da_bindings::OPC_LEAF,
            BrowseType::Flat => opc_da_bindings::OPC_FLAT,
        }
    }
}

impl TryFrom<bindings::tagOPCITEMVQT> for OptionalVqt {
    type Error = windows_core::Error;

    fn try_from(value: bindings::tagOPCITEMVQT) -> Result<Self, Self::Error> {
        Ok(OptionalVqt {
            value: (*value.vDataValue).clone().into(),
            quality: if value.bQualitySpecified.as_bool() {
                Some(value.wQuality)
            } else {
                None
            },
            timestamp: if value.bTimeStampSpecified.as_bool() {
                Some(value.ftTimeStamp.into())
            } else {
                None
            },
        })
    }
}
