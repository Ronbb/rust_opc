use crate::{
    client::{v1, v2, v3, ItemMgtTrait},
    def::{IntoBridge as _, ItemDef, ItemResult, TryFromNative as _, TryToNative as _},
};

pub enum Group {
    V1(v1::Group),
    V2(v2::Group),
    V3(v3::Group),
}

impl Group {
    #[inline(always)]
    fn item_mgt(&self) -> &dyn ItemMgtTrait {
        match self {
            Self::V1(group) => group,
            Self::V2(group) => group,
            Self::V3(group) => group,
        }
    }

    pub fn add_items(
        &self,
        items: Vec<ItemDef>,
    ) -> windows::core::Result<Vec<windows::core::Result<ItemResult>>> {
        let bridge = items.into_bridge();
        let (results, errors) = self.item_mgt().add_items(&bridge.try_to_native()?)?;
        Ok(results
            .as_slice()
            .iter()
            .zip(errors.as_slice())
            .map(|(result, error)| {
                if error.is_ok() {
                    ItemResult::try_from_native(result)
                } else {
                    Err((*error).into())
                }
            })
            .collect())
    }
}

impl From<v1::Group> for Group {
    fn from(group: v1::Group) -> Self {
        Self::V1(group)
    }
}

impl From<v2::Group> for Group {
    fn from(group: v2::Group) -> Self {
        Self::V2(group)
    }
}

impl From<v3::Group> for Group {
    fn from(group: v3::Group) -> Self {
        Self::V3(group)
    }
}
