use crate::{
    client::{v1, v2, v3, ItemMgtTrait, SyncIo2Trait, SyncIoTrait},
    def::{DataSourceTarget, ItemDef, ItemResult, ItemState, ItemValue},
    utils::{IntoBridge as _, TryToLocal as _, TryToNative as _},
};

pub struct Group {
    inner: GroupInner,
    items: Vec<Item>,
}

pub enum GroupInner {
    V1(v1::Group),
    V2(v2::Group),
    V3(v3::Group),
}

pub struct Item {
    pub name: String,
    pub server_handle: u32,
    pub client_handle: u32,
}

impl Group {
    fn new_with_inner(inner: GroupInner) -> Self {
        Self {
            inner,
            items: Vec::new(),
        }
    }

    pub fn items(&self) -> &[Item] {
        &self.items
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

    fn read_items_sync<T: SyncIoTrait>(
        &self,
        sync_io: &T,
        data_source: DataSourceTarget,
        server_handles: &[u32],
    ) -> windows::core::Result<Vec<windows::core::Result<ItemValue>>> {
        let results: Vec<windows::core::Result<ItemState>> = sync_io
            .read(data_source.try_to_native()?, server_handles)?
            .try_to_local()?;

        Ok(results
            .into_iter()
            .map(|r| {
                r.map(|r| ItemValue {
                    value: r.data_value,
                    quality: r.quality,
                    timestamp: r.timestamp,
                })
            })
            .collect())
    }

    fn read_items_sync2<T: SyncIo2Trait>(
        &self,
        sync_io2: &T,
        server_handles: &[u32],
        max_ages: &[u32],
    ) -> windows::core::Result<Vec<windows::core::Result<ItemValue>>> {
        sync_io2
            .read_max_age(server_handles, max_ages)?
            .try_to_local()
    }

    pub fn read_items<S>(
        &self,
        items_names: &[S],
        data_source: DataSourceTarget,
    ) -> windows::core::Result<Vec<windows::core::Result<ItemValue>>>
    where
        S: AsRef<str>,
    {
        let server_handles: Vec<u32> = items_names
            .iter()
            .map(|name| {
                self.items
                    .iter()
                    .find(|item| item.name == name.as_ref())
                    .map(|item| item.server_handle)
                    .ok_or_else(|| {
                        windows::core::Error::new(
                            windows::Win32::Foundation::E_INVALIDARG,
                            "item name not found",
                        )
                    })
            })
            .collect::<windows::core::Result<_>>()?;

        match &self.inner {
            GroupInner::V1(group) => self.read_items_sync(group, data_source, &server_handles),
            GroupInner::V2(group) => self.read_items_sync(group, data_source, &server_handles),
            GroupInner::V3(group) => self.read_items_sync2(
                group,
                &server_handles,
                &vec![data_source.max_age(); server_handles.len()],
            ),
        }
    }
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
