use std::{collections::BTreeMap, mem::ManuallyDrop, sync::Arc};

use tokio::sync::Mutex;
use windows::Win32::{
    Foundation::{S_FALSE, S_OK},
    System::Com::{
        IConnectionPoint, IEnumConnectionPoints, IEnumConnectionPoints_Impl, IEnumConnections,
        IEnumConnections_Impl, IEnumString, IEnumString_Impl, IEnumUnknown, IEnumUnknown_Impl,
        CONNECTDATA,
    },
};
use windows_core::{implement, IUnknown};

use super::{
    bindings,
    utils::{PointerWriter, TryWriteTo},
};

#[implement(IEnumString)]
pub struct StringEnumerator {
    strings: Arc<Vec<String>>,
    index: Mutex<usize>,
}

#[implement(IEnumUnknown)]
pub struct UnknownEnumerator {
    items: Arc<Vec<IUnknown>>,
    index: Mutex<usize>,
}

#[implement(IEnumConnectionPoints)]
pub struct ConnectionPointsEnumerator {
    pub connection_points: Vec<IConnectionPoint>,
    index: Mutex<usize>,
}

#[implement(IEnumConnections)]
pub struct ConnectionsEnumerator {
    pub connections: Arc<Vec<CONNECTDATA>>,
    index: Mutex<usize>,
}

#[implement(bindings::IEnumOPCItemAttributes)]
pub struct ItemAttributesEnumerator {
    index: Mutex<usize>,
    items: Arc<Vec<bindings::tagOPCITEMATTRIBUTES>>,
}

impl StringEnumerator {
    pub fn new(strings: Vec<String>) -> Self {
        Self {
            strings: Arc::new(strings),
            index: Mutex::new(0),
        }
    }
}

impl UnknownEnumerator {
    pub fn new(items: Vec<IUnknown>) -> Self {
        Self {
            items: Arc::new(items),
            index: Mutex::new(0),
        }
    }
}

impl ConnectionPointsEnumerator {
    pub fn new(connection_points: Vec<IConnectionPoint>) -> Self {
        Self {
            connection_points,
            index: Mutex::new(0),
        }
    }
}

impl ConnectionsEnumerator {
    pub fn new(connections: Arc<Vec<CONNECTDATA>>) -> Self {
        Self {
            connections,
            index: Mutex::new(0),
        }
    }

    pub fn from_map(map: BTreeMap<u32, windows_core::IUnknown>) -> Self {
        let connections = map
            .into_iter()
            .map(|(cookie, unknown)| CONNECTDATA {
                dwCookie: cookie,
                pUnk: ManuallyDrop::new(Some(unknown)),
            })
            .collect();
        Self {
            connections: Arc::new(connections),
            index: Mutex::new(0),
        }
    }
}

impl ItemAttributesEnumerator {
    pub fn new(items: Vec<bindings::tagOPCITEMATTRIBUTES>) -> Self {
        Self {
            index: Mutex::new(0),
            items: Arc::new(items),
        }
    }
}

impl IEnumString_Impl for StringEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_elements: *mut windows_core::PWSTR,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let mut index = self.index.blocking_lock();
        if *index >= self.strings.len() {
            unsafe { *count_fetched = 0 };
            S_FALSE
        } else {
            let mut fetched = 0;
            while fetched < count && *index < self.strings.len() {
                let buffer: windows_core::Result<_> =
                    PointerWriter::try_write_to(&self.strings[*index]);
                let buffer = match buffer {
                    Ok(buffer) => buffer,
                    Err(e) => return e.code(),
                };

                unsafe { range_elements.add(fetched as usize).write(buffer) };
                fetched += 1;
                *index += 1;
            }
            unsafe { *count_fetched = fetched };
            S_OK
        }
    }

    fn Skip(&self, count: u32) -> windows_core::HRESULT {
        let mut index = self.index.blocking_lock();
        if *index + count as usize > self.strings.len() {
            *index = self.strings.len();
            S_FALSE
        } else {
            *index += count as usize;
            S_OK
        }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        *index = 0;
        Ok(())
    }

    fn Clone(&self) -> windows_core::Result<IEnumString> {
        Ok(IEnumString::from(StringEnumerator {
            strings: self.strings.clone(),
            index: Mutex::new(self.index.blocking_lock().clone()),
        }))
    }
}

impl IEnumUnknown_Impl for UnknownEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_elements: *mut Option<windows_core::IUnknown>,
        fetched_count: *mut u32,
    ) -> windows_core::HRESULT {
        let mut index = self.index.blocking_lock();
        if *index >= self.items.len() {
            unsafe { *fetched_count = 0 };
            return S_FALSE;
        }

        let mut fetched = 0;
        while fetched < count && *index < self.items.len() {
            unsafe {
                range_elements
                    .add(fetched as usize)
                    .write(Some(self.items[*index].clone()));
            }
            fetched += 1;
            *index += 1;
        }
        unsafe { *fetched_count = fetched };
        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        if *index + count as usize > self.items.len() {
            *index = self.items.len();
            Ok(())
        } else {
            *index += count as usize;
            Ok(())
        }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        *index = 0;
        Ok(())
    }

    fn Clone(&self) -> windows_core::Result<IEnumUnknown> {
        Ok(IEnumUnknown::from(UnknownEnumerator {
            items: self.items.clone(),
            index: Mutex::new(self.index.blocking_lock().clone()),
        }))
    }
}

impl IEnumConnectionPoints_Impl for ConnectionPointsEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_connection_points: *mut Option<IConnectionPoint>,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let mut index = self.index.blocking_lock();
        if *index >= self.connection_points.len() {
            unsafe { *count_fetched = 0 };
            return S_FALSE;
        }

        let mut fetched = 0;
        while fetched < count && *index < self.connection_points.len() {
            unsafe {
                range_connection_points
                    .add(fetched as usize)
                    .write(Some(self.connection_points[*index].clone()));
            }
            fetched += 1;
            *index += 1;
        }
        unsafe { *count_fetched = fetched };
        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        if *index + count as usize > self.connection_points.len() {
            *index = self.connection_points.len();
            Ok(())
        } else {
            *index += count as usize;
            Ok(())
        }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        *index = 0;
        Ok(())
    }

    fn Clone(&self) -> windows_core::Result<IEnumConnectionPoints> {
        Ok(IEnumConnectionPoints::from(ConnectionPointsEnumerator {
            connection_points: self.connection_points.clone(),
            index: Mutex::new(self.index.blocking_lock().clone()),
        }))
    }
}

impl IEnumConnections_Impl for ConnectionsEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_connect_data: *mut CONNECTDATA,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let mut index = self.index.blocking_lock();
        if *index >= self.connections.len() {
            unsafe { *count_fetched = 0 };
            return S_FALSE;
        }

        let mut fetched = 0;
        while fetched < count && *index < self.connections.len() {
            unsafe {
                range_connect_data
                    .add(fetched as usize)
                    .write(self.connections[*index].clone());
            }
            fetched += 1;
            *index += 1;
        }
        unsafe { *count_fetched = fetched };
        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        if *index + count as usize > self.connections.len() {
            *index = self.connections.len();
            Ok(())
        } else {
            *index += count as usize;
            Ok(())
        }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        *index = 0;
        Ok(())
    }

    fn Clone(&self) -> windows_core::Result<IEnumConnections> {
        Ok(IEnumConnections::from(ConnectionsEnumerator {
            connections: self.connections.clone(),
            index: Mutex::new(self.index.blocking_lock().clone()),
        }))
    }
}

impl bindings::IEnumOPCItemAttributes_Impl for ItemAttributesEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        items: *mut *mut bindings::tagOPCITEMATTRIBUTES,
        fetched_count: *mut u32,
    ) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        if *index >= self.items.len() {
            unsafe { *fetched_count = 0 };
            return Ok(());
        }

        let items: *mut *const bindings::tagOPCITEMATTRIBUTES = items.cast();

        let mut fetched = 0;
        while fetched < count && *index < self.items.len() {
            unsafe {
                items.add(fetched as usize).write(&self.items[*index]);
            }
            fetched += 1;
            *index += 1;
        }
        unsafe { *fetched_count = fetched };
        Ok(())
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        if *index + count as usize > self.items.len() {
            *index = self.items.len();
        } else {
            *index += count as usize;
        }
        Ok(())
    }

    fn Reset(&self) -> windows_core::Result<()> {
        let mut index = self.index.blocking_lock();
        *index = 0;
        Ok(())
    }

    fn Clone(&self) -> windows_core::Result<bindings::IEnumOPCItemAttributes> {
        Ok(bindings::IEnumOPCItemAttributes::from(
            ItemAttributesEnumerator {
                index: Mutex::new(self.index.blocking_lock().clone()),
                items: self.items.clone(),
            },
        ))
    }
}
