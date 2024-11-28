use std::{
    collections::BTreeMap,
    ops::Deref,
    sync::{Arc, Weak},
};

use crate::traits::ServerTrait;

use super::{
    base::{Core, Quality, Variant},
    utils::copy_to_pointer,
};

use super::{
    bindings::{self, tagOPCITEMPROPERTIES, IOPCShutdown},
    connection_point::ConnectionPoint,
    enumeration::{ConnectionPointsEnumerator, StringEnumerator, UnknownEnumerator},
    group::Group,
    utils::{com_alloc_v, copy_to_com_string},
};
use globset::Glob;
use windows::Win32::{
    Foundation::{E_INVALIDARG, E_NOTIMPL, E_POINTER, S_OK},
    System::Com::{
        IConnectionPoint, IConnectionPointContainer, IConnectionPointContainer_Impl,
        IEnumConnectionPoints, IEnumString, IEnumUnknown,
    },
};
use windows_core::{
    implement, w, ComObject, ComObjectInner, IUnknownImpl, Interface, PWSTR, VARIANT,
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
        _locale: u32,
    ) -> windows_core::Result<windows_core::PWSTR> {
        let message = error.message();
        Ok(copy_to_com_string(&message))
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
        copy_to_pointer(&available_locale_ids, count, locale_ids)?;
        Ok(())
    }

    fn GetErrorString(
        &self,
        error: windows_core::HRESULT,
    ) -> windows_core::Result<windows_core::PWSTR> {
        let s = self.get_error_string(error.0)?;
        copy_to_com_string(&s)
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
        Ok(
            ConnectionPointsEnumerator::new(Arc::new(vec![self.shutdown.to_interface()]))
                .into_object()
                .into_interface(),
        )
    }

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<IConnectionPoint> {
        match unsafe { reference_interface_id.as_ref() } {
            Some(&IOPCShutdown::IID) => Ok(self.shutdown.to_interface()),
            _ => Err(E_INVALIDARG.into()),
        }
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
        let node = tokio::runtime::Handle::current().block_on(
            self.core
                .get_node_from_path(&unsafe { item_id.to_string() }?),
        );
        match node {
            Some(node) => {
                let node = node.blocking_read();
                let item_properties = node.get_item_properties_without_value();
                let mut i = Vec::with_capacity(item_properties.len());
                let mut d = Vec::with_capacity(item_properties.len());
                let mut t = Vec::with_capacity(item_properties.len());

                for property in item_properties.iter() {
                    i.push(property.id);
                    d.push(copy_to_com_string(&property.description));
                    t.push(property.value.get_data_type());
                }

                unsafe {
                    *count = i.len() as u32;
                    *property_ids = com_alloc_v(&i);
                    *descriptions = com_alloc_v(&d);
                    *data_types = com_alloc_v(&t);
                }

                Ok(())
            }
            None => Err(E_INVALIDARG.into()),
        }
    }

    fn GetItemProperties(
        &self,
        item_id: &windows_core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        data: *mut *mut windows_core::VARIANT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let node = tokio::runtime::Handle::current().block_on(
            self.core
                .get_node_from_path(&unsafe { item_id.to_string() }?),
        );
        match node {
            Some(node) => {
                let node = node.blocking_read();
                let count = count as usize;
                let mut v = Vec::with_capacity(count);
                let mut e = Vec::with_capacity(count);

                for i in 0..count {
                    let property_id = unsafe { *property_ids.add(i) };
                    if let Some(property) = node.get_item_property(property_id) {
                        v.push(property.value.clone().into());
                        e.push(S_OK);
                    } else {
                        v.push(VARIANT::default());
                        e.push(E_INVALIDARG);
                    }
                }

                unsafe {
                    *data = com_alloc_v(&v);
                    *errors = com_alloc_v(&e);
                }

                Ok(())
            }
            None => Err(E_INVALIDARG.into()),
        }
    }

    fn LookupItemIDs(
        &self,
        item_id: &windows_core::PCWSTR,
        count: u32,
        property_ids: *const u32,
        new_item_ids: *mut *mut windows_core::PWSTR,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let item_id = unsafe { item_id.to_string() }?;
        let node =
            tokio::runtime::Handle::current().block_on(self.core.get_node_from_path(&item_id));
        match node {
            Some(node) => {
                let node = node.blocking_read();
                let count = count as usize;
                let mut n = Vec::with_capacity(count);
                let mut e = Vec::with_capacity(count);

                for i in 0..count {
                    let property_id = unsafe { *property_ids.add(i) };
                    if let Some(property) = node.get_item_property(property_id) {
                        n.push(copy_to_com_string(&format!(
                            "{}{}{}",
                            item_id, self.seperator, property.name
                        )));
                        e.push(S_OK);
                    } else {
                        n.push(PWSTR::null());
                        e.push(E_INVALIDARG);
                    }
                }

                unsafe {
                    *new_item_ids = com_alloc_v(&n);
                    *errors = com_alloc_v(&e);
                }

                Ok(())
            }
            None => Err(E_INVALIDARG.into()),
        }
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
        let item_count = item_count as usize;
        let property_count = property_count as usize;
        let mut properties = Vec::with_capacity(item_count);
        for i in 0..item_count {
            let item_id = unsafe { (*item_ids.add(i)).to_string()? };
            let node =
                tokio::runtime::Handle::current().block_on(self.core.get_node_from_path(&item_id));
            match node {
                Some(node) => {
                    let node = node.blocking_read();
                    let mut ps = Vec::with_capacity(property_count);
                    for j in 0..property_count {
                        let property_id = unsafe { *property_ids.add(j) };
                        let property = match return_property_values {
                            windows::Win32::Foundation::FALSE => {
                                node.get_item_property_without_value(property_id)
                            }
                            windows::Win32::Foundation::TRUE => node.get_item_property(property_id),
                            _ => return Err(E_INVALIDARG.into()),
                        };

                        if let Some(property) = property {
                            ps.push(property.into());
                        }
                    }

                    properties.push(tagOPCITEMPROPERTIES {
                        hrErrorID: S_OK,
                        dwNumProperties: property_count as u32,
                        pItemProperties: com_alloc_v(&ps),
                        dwReserved: 0,
                    });
                }
                None => {
                    properties.push(tagOPCITEMPROPERTIES::default());
                }
            }
        }

        unsafe { *item_properties = com_alloc_v(&properties) };

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
        let item_id = if item_id.is_null() {
            "".to_string()
        } else {
            unsafe { item_id.to_string()? }
        };

        let continuation_point = if continuation_point.is_null() {
            "".to_string()
        } else {
            unsafe { continuation_point.read().to_string()? }
        };

        if browse_filter.0 < bindings::OPC_BROWSE_FILTER_ALL.0
            || browse_filter.0 > bindings::OPC_BROWSE_FILTER_ITEMS.0
        {
            return Err(E_INVALIDARG.into());
        }

        unsafe { more_elements.write(windows::Win32::Foundation::FALSE) };

        let browse_root =
            tokio::runtime::Handle::current().block_on(self.core.get_node_from_path(&item_id));
        match browse_root {
            Some(browse_root) => {
                let browse_root = browse_root.blocking_read();
                let mut elements = Vec::new();

                let children = browse_root.children.blocking_read();
                for (name, child) in children.range(continuation_point..) {
                    if !element_name_filter.is_null() {
                        if !Glob::new(unsafe { &element_name_filter.to_string()? })
                            .map_err(|_| E_INVALIDARG)?
                            .compile_matcher()
                            .is_match(&name)
                        {
                            continue;
                        }
                    }

                    if !vendor_filter.is_null() {
                        if !Glob::new(unsafe { &vendor_filter.to_string()? })
                            .map_err(|_| E_INVALIDARG)?
                            .compile_matcher()
                            .is_match(&name)
                        {
                            continue;
                        }
                    }

                    if max_elements_returned > 0 && elements.len() >= max_elements_returned as usize
                    {
                        unsafe { more_elements.write(windows::Win32::Foundation::TRUE) };
                        break;
                    }

                    let child = child.blocking_read();

                    match browse_filter {
                        bindings::OPC_BROWSE_FILTER_ALL => {}
                        bindings::OPC_BROWSE_FILTER_BRANCHES => {
                            if !child.children.blocking_read().is_empty() {
                                continue;
                            }
                        }
                        bindings::OPC_BROWSE_FILTER_ITEMS => {
                            if child.children.blocking_read().is_empty() {
                                continue;
                            }
                        }
                        _ => {}
                    }

                    let mut properties = Vec::new();
                    if return_all_properties.as_bool() {
                        let item_properties = child.get_item_properties_without_value();
                        for property in item_properties.iter() {
                            properties.push(property.clone().into());
                        }
                    } else {
                        for i in 0..property_count {
                            let property_id = unsafe { *property_ids.add(i as usize) };
                            let property = match return_property_values {
                                windows::Win32::Foundation::FALSE => {
                                    child.get_item_property_without_value(property_id)
                                }
                                windows::Win32::Foundation::TRUE => {
                                    child.get_item_property(property_id)
                                }
                                _ => return Err(E_INVALIDARG.into()),
                            };

                            if let Some(property) = property {
                                properties.push(property.into());
                            }
                        }
                    }

                    elements.push(bindings::tagOPCBROWSEELEMENT {
                        szName: copy_to_com_string(&child.name),
                        szItemID: copy_to_com_string(
                            &tokio::runtime::Handle::current().block_on(child.get_path()),
                        ),
                        dwFlagValue: 0,
                        dwReserved: 0,
                        ItemProperties: tagOPCITEMPROPERTIES {
                            hrErrorID: S_OK,
                            dwNumProperties: properties.len() as u32,
                            pItemProperties: com_alloc_v(&properties),
                            dwReserved: 0,
                        },
                    });
                }

                unsafe {
                    *count = elements.len() as u32;
                    *browse_elements = com_alloc_v(&elements);
                }

                Ok(())
            }
            None => Err(E_INVALIDARG.into()),
        }
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
        match self.public_server.upgrade() {
            Some(server) => {
                let groups = server.group_name_map.blocking_read();
                if let Some(group) = groups.get(&unsafe { name.to_string()? }) {
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

                    Ok(unknown.unwrap())
                } else {
                    Err(E_INVALIDARG.into())
                }
            }
            None => Err(E_INVALIDARG.into()),
        }
    }

    fn RemovePublicGroup(
        &self,
        server_group: u32,
        _force: windows::Win32::Foundation::BOOL,
    ) -> windows_core::Result<()> {
        match self.public_server.upgrade() {
            Some(server) => {
                let group = server
                    .group_server_handle_map
                    .blocking_write()
                    .remove(&server_group);
                if let Some(group) = group {
                    let mut groups = server.group_name_map.blocking_write();
                    groups.remove(&group.state.blocking_read().name);
                    Ok(())
                } else {
                    Err(E_INVALIDARG.into())
                }
            }
            None => Err(E_INVALIDARG.into()),
        }
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: ServerTrait + 'static> bindings::IOPCBrowseServerAddressSpace_Impl for Server_Impl<T> {
    fn QueryOrganization(&self) -> windows_core::Result<bindings::tagOPCNAMESPACETYPE> {
        Ok(bindings::OPC_NS_HIERARCHIAL)
    }

    fn ChangeBrowsePosition(
        &self,
        browse_direction: bindings::tagOPCBROWSEDIRECTION,
        string: &windows_core::PCWSTR,
    ) -> windows_core::Result<()> {
        let mut position = self.browser_position.blocking_write();

        match browse_direction {
            bindings::OPC_BROWSE_DOWN => {
                *position = format!("{}{}{}", position, self.seperator, unsafe {
                    string.to_string().unwrap()
                });
            }
            bindings::OPC_BROWSE_TO => {
                *position = unsafe { string.to_string().unwrap() };
            }
            bindings::OPC_BROWSE_UP => {
                if let Some(index) = position.rfind(&self.seperator) {
                    position.truncate(index);
                } else {
                    position.clear();
                }
            }
            _ => return Err(E_INVALIDARG.into()),
        };

        Ok(())
    }

    fn BrowseOPCItemIDs(
        &self,
        browse_filter_type: bindings::tagOPCBROWSETYPE,
        filter_criteria: &windows_core::PCWSTR,
        variant_data_type_filter: u16,
        access_rights_filter: u32,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumString> {
        let filter_criteria = unsafe { filter_criteria.to_string()? };
        let filter_criteria = if filter_criteria.is_empty() {
            None
        } else {
            Some(Glob::new(&filter_criteria).map_err(|_| E_INVALIDARG)?)
        };

        let variant_data_type_filter = if variant_data_type_filter == 0 {
            None
        } else {
            Some(variant_data_type_filter)
        };

        let access_rights_filter = if access_rights_filter == 0 {
            None
        } else {
            Some(access_rights_filter)
        };

        tokio::runtime::Handle::current().block_on(async {
            let node = self
                .core
                .get_node_from_path(&self.browser_position.read().await.clone())
                .await;

            match node {
                Some(node) => {
                    let node = node.read().await;
                    let mut items = vec![];
                    let children = node.children.read().await;
                    for (name, child) in children.iter() {
                        if let Some(filter_criteria) = &filter_criteria {
                            if !filter_criteria.compile_matcher().is_match(name) {
                                continue;
                            }
                        }

                        let child = child.read().await;

                        match browse_filter_type {
                            bindings::OPC_BRANCH => {
                                if child.children.read().await.is_empty() {
                                    continue;
                                }
                            }
                            bindings::OPC_LEAF => {
                                if !child.children.read().await.is_empty() {
                                    continue;
                                }
                            }
                            bindings::OPC_FLAT => {}
                            _ => {
                                return Err(E_INVALIDARG.into());
                            }
                        }

                        if let Some(variant_data_type_filter) = variant_data_type_filter {
                            if variant_data_type_filter != Variant::Empty.get_data_type() {
                                let data_type = child.value.read().await.variant.get_data_type();
                                if data_type != variant_data_type_filter {
                                    continue;
                                }
                            }
                        }

                        if let Some(access_rights_filter) = access_rights_filter {
                            if access_rights_filter != 0 {
                                let access_right = child.access_right.read().await;
                                if (access_rights_filter & bindings::OPC_READABLE) != 0
                                    && !access_right.readable
                                {
                                    continue;
                                }
                                if (access_rights_filter & bindings::OPC_WRITEABLE) != 0
                                    && !access_right.writable
                                {
                                    continue;
                                }
                            }
                        }

                        items.push(name.clone());
                    }

                    Ok(StringEnumerator::new(items).into_object().into_interface())
                }
                None => Err(E_INVALIDARG.into()),
            }
        })
    }

    fn GetItemID(
        &self,
        item_data_id: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows_core::PWSTR> {
        if item_data_id.is_null() {
            return Err(E_POINTER.into());
        }

        let item_data_id = unsafe { item_data_id.to_string()? };

        tokio::runtime::Handle::current().block_on(async {
            let node = self
                .core
                .get_node_from_path(&format!(
                    "{}{}{}",
                    self.browser_position.read().await,
                    self.seperator,
                    item_data_id
                ))
                .await;

            match node {
                Some(node) => {
                    let path = node.read().await.get_path().await;
                    Ok(copy_to_com_string(&path))
                }
                None => Err(E_INVALIDARG.into()),
            }
        })
    }

    fn BrowseAccessPaths(
        &self,
        _item_id: &windows_core::PCWSTR,
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumString> {
        Err(E_NOTIMPL.into())
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
        _max_age: *const u32,
        values: *mut *mut windows_core::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let count = count as usize;
        let mut v = Vec::with_capacity(count);
        let mut q = Vec::with_capacity(count);
        let mut t = Vec::with_capacity(count);
        let mut e = Vec::with_capacity(count);

        for i in 0..count {
            let item_id = unsafe { (*item_ids.add(i)).to_string()? };
            let node =
                tokio::runtime::Handle::current().block_on(self.core.get_node_from_path(&item_id));
            match node {
                Some(node) => {
                    let node = node.blocking_read();
                    let value = node.value.blocking_read();

                    v.push(value.variant.clone().into());
                    q.push(value.quality.to_u16());
                    t.push(match &value.timestamp {
                        Some(timestamp) => timestamp.clone().into(),
                        None => windows::Win32::Foundation::FILETIME::default(),
                    });
                    e.push(S_OK);
                }
                None => {
                    v.push(VARIANT::default());
                    q.push(0);
                    t.push(windows::Win32::Foundation::FILETIME::default());
                    e.push(E_INVALIDARG);
                }
            }
        }

        unsafe {
            *values = com_alloc_v(&v);
            *qualities = com_alloc_v(&q);
            *timestamps = com_alloc_v(&t);
            *errors = com_alloc_v(&e);
        }

        Ok(())
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_ids: *const windows_core::PCWSTR,
        item_vqt: *const bindings::tagOPCITEMVQT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let count = count as usize;
        let mut e = Vec::with_capacity(count);

        for i in 0..count {
            let item_id = unsafe { (*item_ids.add(i)).to_string()? };
            let node =
                tokio::runtime::Handle::current().block_on(self.core.get_node_from_path(&item_id));
            match node {
                Some(node) => {
                    let node = node.blocking_read();
                    let mut value = node.value.blocking_write();
                    let item_vqt = unsafe { item_vqt.add(i) };
                    value.variant = unsafe { item_vqt.read() }
                        .vDataValue
                        .as_raw()
                        .clone()
                        .into();
                    value.quality = Quality(unsafe { item_vqt.read().wQuality });
                    value.timestamp = match unsafe { item_vqt.read().ftTimeStamp } {
                        windows::Win32::Foundation::FILETIME {
                            dwLowDateTime,
                            dwHighDateTime,
                        } => {
                            if dwLowDateTime == 0 && dwHighDateTime == 0 {
                                None
                            } else {
                                Some(unsafe { (*item_vqt).ftTimeStamp }.into())
                            }
                        }
                    };
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }

        unsafe {
            *errors = com_alloc_v(&e);
        }

        Ok(())
    }
}
