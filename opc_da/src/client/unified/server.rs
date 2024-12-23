use crate::{
    client::{v1, v2, v3, ServerTrait},
    def::{BrowseFilter, BrowseType, EnumScope, GroupState, ServerStatus},
    utils::{ToNative as _, TryToLocal},
};

use super::Group;

pub enum Server {
    V1(v1::Server),
    V2(v2::Server),
    V3(v3::Server),
}

impl Server {
    fn add_group_with_server<
        G: TryFrom<windows::core::IUnknown, Error = windows::core::Error>,
        T: ServerTrait<G>,
    >(
        server: &T,
        mut state: GroupState,
    ) -> windows::core::Result<G> {
        server.add_group(
            &state.name,
            state.active,
            state.client_handle,
            state.update_rate,
            state.locale_id,
            state.time_bias,
            state.percent_deadband,
            &mut state.update_rate,
            &mut state.server_handle,
        )
    }

    pub fn add_group(&self, state: GroupState) -> windows::core::Result<Group> {
        match self {
            Self::V1(server) => Ok(Self::add_group_with_server(server, state)?.into()),
            Self::V2(server) => Ok(Self::add_group_with_server(server, state)?.into()),
            Self::V3(server) => Ok(Self::add_group_with_server(server, state)?.into()),
        }
    }

    pub fn get_status(&self) -> windows::core::Result<ServerStatus> {
        let status = match self {
            Self::V1(server) => server.get_status(),
            Self::V2(server) => server.get_status(),
            Self::V3(server) => server.get_status(),
        }?;

        status.ok()?.try_to_local()
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
        scope: EnumScope,
    ) -> windows::core::Result<GroupIterator> {
        let scope = scope.to_native();

        let iterator = match self {
            Self::V1(server) => GroupIterator::V1(server.create_group_enumerator(scope)?),
            Self::V2(server) => GroupIterator::V2(server.create_group_enumerator(scope)?),
            Self::V3(server) => GroupIterator::V3(server.create_group_enumerator(scope)?),
        };

        Ok(iterator)
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
    pub browse_type: BrowseType,
    pub browse_filter: BrowseFilter,
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
