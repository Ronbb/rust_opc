use std::{collections::BTreeMap, sync::Arc};

use tokio::sync::RwLock;

use super::variant::Variant;

#[derive(Default)]
pub struct Node {
    pub name: String,
    pub value: RwLock<Value>,
    pub children: RwLock<BTreeMap<String, Arc<RwLock<Node>>>>,
    pub parent: Option<Arc<RwLock<Node>>>,
    pub access_right: RwLock<AccessRight>,
    pub state: RwLock<NodeState>,
}

#[derive(Default)]
pub struct NodeState {
    pub is_active: bool,
}

#[derive(Clone, Default)]
pub struct Quality(pub u16);

#[derive(Clone)]
pub struct SystemTime(pub std::time::SystemTime);

#[derive(Default)]
pub struct Value {
    pub variant: Variant,
    pub quality: Quality,
    pub timestamp: Option<SystemTime>,
}

#[derive(Default)]
pub struct AccessRight {
    pub readable: bool,
    pub writable: bool,
}

#[derive(Default)]
pub struct Core {
    root: Arc<RwLock<Node>>,
}

impl Core {
    pub fn new() -> Self {
        Core::default()
    }

    pub fn root(&self) -> Arc<RwLock<Node>> {
        self.root.clone()
    }

    pub async fn get_node_from_path(&self, path: &String) -> Option<Arc<RwLock<Node>>> {
        self.root.read().await.get_node_from_path(path).await
    }
}

impl Node {
    pub async fn get_path(&self) -> String {
        todo!()
    }

    pub async fn get_node_from_path(&self, _path: &String) -> Option<Arc<RwLock<Node>>> {
        todo!()
    }
}
