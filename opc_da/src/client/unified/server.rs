use crate::{
    client::{v1, v2, v3, ServerTrait as _},
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

        def::ServerStatus::try_from_native(status.as_result()?)
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
