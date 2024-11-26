use std::{mem::ManuallyDrop, sync::Arc};

use tokio::sync::RwLock;
use windows::Win32::Foundation::S_OK;

use crate::core::{core::Node, variant::Variant};

use super::{bindings::tagOPCITEMPROPERTY, utils::com_alloc_str};

pub struct Item {
    pub name: String,
    pub server_handle: u32,
    pub client_handle: u32,
    pub node: Arc<RwLock<Node>>,
}

#[derive(Clone)]
pub struct ItemProperty {
    pub id: u32,
    pub name: String,
    pub description: String,
    pub value: Variant,
}

impl Node {
    pub fn get_item_properties(&self) -> Vec<ItemProperty> {
        todo!()
    }

    pub fn get_item_properties_without_value(&self) -> Vec<ItemProperty> {
        todo!()
    }

    pub fn get_item_property(&self, _id: u32) -> Option<ItemProperty> {
        todo!()
    }

    pub fn get_item_property_without_value(&self, _id: u32) -> Option<ItemProperty> {
        todo!()
    }
}

impl Into<tagOPCITEMPROPERTY> for ItemProperty {
    fn into(self) -> tagOPCITEMPROPERTY {
        tagOPCITEMPROPERTY {
            vtDataType: self.value.get_data_type(),
            wReserved: 0,
            dwPropertyID: self.id,
            szItemID: com_alloc_str(&self.name),
            szDescription: com_alloc_str(&self.description),
            vValue: ManuallyDrop::new(self.value.into()),
            hrErrorID: S_OK,
            dwReserved: 0,
        }
    }
}
