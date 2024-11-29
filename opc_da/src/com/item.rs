use std::{mem::ManuallyDrop, sync::Arc};

use tokio::sync::RwLock;
use windows::Win32::Foundation::S_OK;

use super::{
    base::{Node, Variant},
    bindings::tagOPCITEMPROPERTY,
    utils::{PointerWriter, TryWriteTo},
};

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

impl TryFrom<ItemProperty> for tagOPCITEMPROPERTY {
    type Error = windows_core::Error;

    fn try_from(val: ItemProperty) -> Result<Self, Self::Error> {
        Ok(tagOPCITEMPROPERTY {
            vtDataType: val.value.get_data_type(),
            wReserved: 0,
            dwPropertyID: val.id,
            szItemID: PointerWriter::try_write_to(&val.name)?,
            szDescription: PointerWriter::try_write_to(&val.description)?,
            vValue: ManuallyDrop::new(val.value.into()),
            hrErrorID: S_OK,
            dwReserved: 0,
        })
    }
}
