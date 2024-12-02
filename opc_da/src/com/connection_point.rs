use std::{collections::BTreeMap, mem::ManuallyDrop};

use windows::Win32::System::Com::{
    IConnectionPoint, IConnectionPointContainer, IConnectionPoint_Impl, IEnumConnections,
};

use super::enumeration::ConnectionsEnumerator;

#[windows_core::implement(IConnectionPoint)]
pub struct ConnectionPoint {
    container: IConnectionPointContainer,
    interface_id: windows_core::GUID,
    next_cookie: core::sync::atomic::AtomicU32,
    connections: tokio::sync::RwLock<BTreeMap<u32, windows_core::IUnknown>>,
}

impl ConnectionPoint {
    pub fn new(
        container: IConnectionPointContainer,
        interface_id: windows_core::GUID,
    ) -> ConnectionPoint {
        ConnectionPoint {
            container,
            interface_id,
            next_cookie: core::sync::atomic::AtomicU32::new(0),
            connections: tokio::sync::RwLock::new(BTreeMap::new()),
        }
    }
}

impl IConnectionPoint_Impl for ConnectionPoint_Impl {
    fn GetConnectionInterface(&self) -> windows_core::Result<windows_core::GUID> {
        Ok(self.interface_id)
    }

    fn GetConnectionPointContainer(&self) -> windows_core::Result<IConnectionPointContainer> {
        Ok(self.container.clone())
    }

    fn Advise(&self, sink: Option<&windows_core::IUnknown>) -> windows_core::Result<u32> {
        let cookie = self
            .next_cookie
            .fetch_add(1, core::sync::atomic::Ordering::SeqCst);
        self.connections
            .blocking_write()
            .insert(cookie, sink.unwrap().clone());
        Ok(cookie)
    }

    fn Unadvise(&self, cookie: u32) -> windows_core::Result<()> {
        self.connections.blocking_write().remove(&cookie);
        Ok(())
    }

    fn EnumConnections(&self) -> windows_core::Result<IEnumConnections> {
        Ok(
            windows_core::ComObjectInner::into_object(ConnectionsEnumerator::new(
                self.connections
                    .blocking_read()
                    .iter()
                    .map(|(k, v)| windows::Win32::System::Com::CONNECTDATA {
                        pUnk: ManuallyDrop::new(Some(v.clone())),
                        dwCookie: *k,
                    })
                    .collect(),
            ))
            .into_interface(),
        )
    }
}
