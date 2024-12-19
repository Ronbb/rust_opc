use crate::{
    client::{v1, v2, v3, ItemMgtTrait},
    def::{IntoBridge as _, ItemDef, ItemResult, TryToLocal as _, TryToNative as _},
};

pub struct Group {
    inner: GroupInner,
}

pub enum GroupInner {
    V1(v1::Group),
    V2(v2::Group),
    V3(v3::Group),
}

struct Item {
    name: String,
    server_handle: u32,
    client_handle: u32,
}

impl Group {
    fn new_with_inner(inner: GroupInner) -> Self {
        Self { inner }
    }
}

impl Group {
    #[inline(always)]
    fn item_mgt(&self) -> &dyn ItemMgtTrait {
        match &self.inner {
            GroupInner::V1(group) => group,
            GroupInner::V2(group) => group,
            GroupInner::V3(group) => group,
        }
    }

    pub fn add_items(
        &self,
        items: Vec<ItemDef>,
    ) -> windows::core::Result<Vec<windows::core::Result<ItemResult>>> {
        let bridge = items.into_bridge();
        self.item_mgt()
            .add_items(&bridge.try_to_native()?)?
            .try_to_local()
    }

    pub fn validate_items(
        &self,
        items: Vec<ItemDef>,
        blob_update: bool,
    ) -> windows::core::Result<Vec<windows::core::Result<ItemResult>>> {
        let bridge = items.into_bridge();
        self.item_mgt()
            .validate_items(&bridge.try_to_native()?, blob_update)?
            .try_to_local()
    }

    pub fn remove_items(
        &self,
        server_handles: Vec<u32>,
    ) -> windows::core::Result<Vec<windows::core::Result<()>>> {
        self.item_mgt()
            .remove_items(&server_handles)?
            .try_to_local()
    }

    // TODO set_active_state
    // TODO set_client_handle
    // TODO set_datatypes
    // TODO create_enumerator
}

impl From<v1::Group> for Group {
    fn from(group: v1::Group) -> Self {
        Self::new_with_inner(GroupInner::V1(group))
    }
}

impl From<v2::Group> for Group {
    fn from(group: v2::Group) -> Self {
        Self::new_with_inner(GroupInner::V2(group))
    }
}

impl From<v3::Group> for Group {
    fn from(group: v3::Group) -> Self {
        Self::new_with_inner(GroupInner::V3(group))
    }
}
