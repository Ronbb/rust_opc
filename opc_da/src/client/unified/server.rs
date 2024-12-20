use crate::{
    client::{
        v1, v2, v3, BrowseServerAddressSpaceTrait as _, BrowseTrait, ServerTrait, StringIterator,
    },
    def::{self, ToNative as _, TryFromNative as _},
};

use super::Group;

pub enum Server {
    V1(v1::Server),
    V2(v2::Server),
    V3(v3::Server),
}

impl Server {
    pub fn add_group(&self, state: def::GroupState) -> windows::core::Result<Group> {
        let mut state = state;

        match self {
            Self::V1(server) => Ok(server
                .add_group(
                    &state.name,
                    state.active,
                    state.client_handle,
                    state.update_rate,
                    state.locale_id,
                    state.time_bias,
                    state.percent_deadband,
                    &mut state.update_rate,
                    &mut state.server_handle,
                )?
                .into()),
            Self::V2(server) => Ok(server
                .add_group(
                    &state.name,
                    state.active,
                    state.client_handle,
                    state.update_rate,
                    state.locale_id,
                    state.time_bias,
                    state.percent_deadband,
                    &mut state.update_rate,
                    &mut state.server_handle,
                )?
                .into()),
            Self::V3(server) => Ok(server
                .add_group(
                    &state.name,
                    state.active,
                    state.client_handle,
                    state.update_rate,
                    state.locale_id,
                    state.time_bias,
                    state.percent_deadband,
                    &mut state.update_rate,
                    &mut state.server_handle,
                )?
                .into()),
        }
    }

    pub fn get_status(&self) -> windows::core::Result<def::ServerStatus> {
        let status = match self {
            Self::V1(server) => server.get_status(),
            Self::V2(server) => server.get_status(),
            Self::V3(server) => server.get_status(),
        }?;

        def::ServerStatus::try_from_native(status.ok()?)
    }

    pub fn remove_group(&self, server_handle: u32, force: bool) -> windows::core::Result<()> {
        match self {
            Self::V1(server) => server.remove_group(server_handle, force),
            Self::V2(server) => server.remove_group(server_handle, force),
            Self::V3(server) => server.remove_group(server_handle, force),
        }
    }

    pub fn create_group_enumerator(
        &self,
        scope: def::EnumScope,
    ) -> windows::core::Result<GroupIterator> {
        let scope = scope.to_native();

        let iterator = match self {
            Self::V1(server) => GroupIterator::V1(server.create_group_enumerator(scope)?),
            Self::V2(server) => GroupIterator::V2(server.create_group_enumerator(scope)?),
            Self::V3(server) => GroupIterator::V3(server.create_group_enumerator(scope)?),
        };

        Ok(iterator)
    }

    pub fn browse_items(
        &self,
        options: BrowseItemsOptions,
    ) -> windows::core::Result<StringIterator> {
        let iterator = match self {
            Self::V1(server) => {
                return Err(windows::core::Error::new(
                    windows::Win32::Foundation::E_NOTIMPL,
                    "Browsing item ids is not implemented in v1",
                ))
            }
            Self::V2(server) => {
                let browse_type = options.browse_type.to_native();
                server.browse_opc_item_ids(
                    browse_type,
                    options.item_id,
                    options.data_type_filter,
                    options.access_rights_filter,
                )?
            }
            Self::V3(server) => {
                server.browse(
                    options.item_id,
                    options.continuation_point,
                    options.max_elements,
                    options.browse_filter.to_native(),
                    options.element_name_filter,
                    options.vendor_filter,
                    options.return_all_properties,
                    options.return_property_values,
                    &options.property_ids,
                )?;

                todo!()
            }
        };

        todo!()
    }
}

impl From<v1::Server> for Server {
    fn from(server: v1::Server) -> Self {
        Self::V1(server)
    }
}

impl From<v2::Server> for Server {
    fn from(server: v2::Server) -> Self {
        Self::V2(server)
    }
}

impl From<v3::Server> for Server {
    fn from(server: v3::Server) -> Self {
        Self::V3(server)
    }
}

pub enum GroupIterator {
    V1(v1::GroupIterator),
    V2(v2::GroupIterator),
    V3(v3::GroupIterator),
}

impl Iterator for GroupIterator {
    type Item = windows::core::Result<Group>;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::V1(iterator) => iterator.next().map(|group| group.map(Group::from)),
            Self::V2(iterator) => iterator.next().map(|group| group.map(Group::from)),
            Self::V3(iterator) => iterator.next().map(|group| group.map(Group::from)),
        }
    }
}

pub struct BrowseItemsOptions {
    pub browse_type: def::BrowseType,
    pub browse_filter: def::BrowseFilter,
    pub item_id: Option<String>,
    pub continuation_point: Option<String>,
    pub data_type_filter: u16,
    pub access_rights_filter: u32,
    pub max_elements: u32,
    pub element_name_filter: Option<String>,
    pub vendor_filter: Option<String>,
    pub return_all_properties: bool,
    pub return_property_values: bool,
    pub property_ids: Vec<u32>,
}
