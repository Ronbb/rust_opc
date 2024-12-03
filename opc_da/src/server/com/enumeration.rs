use windows::Win32::{
    Foundation::{S_FALSE, S_OK},
    System::Com::{
        IConnectionPoint, IEnumConnectionPoints, IEnumConnectionPoints_Impl, IEnumConnections,
        IEnumConnections_Impl, IEnumString, IEnumString_Impl, IEnumUnknown, IEnumUnknown_Impl,
        CONNECTDATA,
    },
};
use windows_core::PWSTR;

use crate::safe_call;

use super::memory::{FreeRaw as _, IntoArrayRef, IntoComArrayRef, IntoRef as _};

struct Enumerator<T> {
    items: Vec<T>,
    index: core::sync::atomic::AtomicUsize,
}

impl<T: Clone> Enumerator<T> {
    fn new(items: Vec<T>) -> Self {
        Self {
            items,
            index: core::sync::atomic::AtomicUsize::default(),
        }
    }

    pub fn next(&self, count: u32, fetched: &mut u32, elements: &mut [T]) -> windows_core::HRESULT {
        let current_index = self
            .index
            .fetch_add(count as _, core::sync::atomic::Ordering::SeqCst);
        if current_index >= self.items.len() {
            return S_FALSE;
        }

        let end_index = (current_index + count as usize).min(self.items.len());
        let slice = &self.items[current_index..end_index];
        *fetched = slice.len() as u32;

        for (i, element) in slice.iter().enumerate() {
            elements[i] = element.clone();
        }

        S_OK
    }

    pub fn skip(&self, count: u32) -> windows_core::HRESULT {
        let current_index = self.index.load(core::sync::atomic::Ordering::SeqCst);
        let new_index = current_index.saturating_add(count as usize);
        let max_index = self.items.len();

        if new_index >= max_index {
            self.index
                .store(max_index, core::sync::atomic::Ordering::SeqCst);
        } else {
            self.index
                .store(new_index, core::sync::atomic::Ordering::SeqCst);
        }

        S_OK
    }

    fn reset(&self) -> windows_core::HRESULT {
        self.index.store(0, core::sync::atomic::Ordering::SeqCst);
        S_OK
    }
}

impl<T: Clone> Clone for Enumerator<T> {
    fn clone(&self) -> Self {
        Self {
            items: self.items.clone(),
            index: core::sync::atomic::AtomicUsize::new(
                self.index.load(core::sync::atomic::Ordering::SeqCst),
            ),
        }
    }
}

#[windows_core::implement(IEnumString)]
#[repr(transparent)]
pub struct StringEnumerator(Enumerator<Vec<u16>>);

#[windows_core::implement(IEnumUnknown)]
pub struct UnknownEnumerator(Enumerator<windows_core::IUnknown>);

#[windows_core::implement(IEnumConnectionPoints)]
pub struct ConnectionPointsEnumerator(Enumerator<IConnectionPoint>);

#[windows_core::implement(IEnumConnections)]
pub struct ConnectionsEnumerator(Enumerator<CONNECTDATA>);

#[windows_core::implement(opc_da_bindings::IEnumOPCItemAttributes)]
pub struct ItemAttributesEnumerator(Enumerator<opc_da_bindings::tagOPCITEMATTRIBUTES>);

impl StringEnumerator {
    pub fn new(strings: Vec<String>) -> Self {
        Self(Enumerator::new(
            strings
                .into_iter()
                .map(|s| s.encode_utf16().chain(Some(0)).collect())
                .collect(),
        ))
    }
}

impl UnknownEnumerator {
    pub fn new(items: Vec<windows_core::IUnknown>) -> Self {
        Self(Enumerator::new(items))
    }
}

impl ConnectionPointsEnumerator {
    pub fn new(connection_points: Vec<IConnectionPoint>) -> Self {
        Self(Enumerator::new(connection_points))
    }
}

impl ConnectionsEnumerator {
    pub fn new(connections: Vec<CONNECTDATA>) -> Self {
        Self(Enumerator::new(connections))
    }
}

impl ItemAttributesEnumerator {
    pub fn new(items: Vec<opc_da_bindings::tagOPCITEMATTRIBUTES>) -> Self {
        Self(Enumerator::new(items))
    }
}

impl IEnumString_Impl for StringEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_elements: *mut windows_core::PWSTR,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let fetched = match count_fetched.into_ref() {
            Ok(fetched) => fetched,
            Err(e) => return e.code(),
        };

        let elements = match range_elements.into_array_ref(count) {
            Ok(elements) => elements,
            Err(e) => return e.code(),
        };

        let mut strings = Vec::with_capacity(count as usize);

        let code = self.0.next(count, fetched, &mut strings);
        if code != S_OK {
            return code;
        }

        for (i, string) in strings.iter_mut().enumerate() {
            let pwstr = PWSTR::from_raw(string.as_mut_ptr());
            elements[i] = pwstr;
        }

        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::HRESULT {
        self.0.skip(count)
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset().ok()
    }

    fn Clone(&self) -> windows_core::Result<IEnumString> {
        Ok(IEnumString::from(StringEnumerator(self.0.clone())))
    }
}

impl IEnumUnknown_Impl for UnknownEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_elements: *mut Option<windows_core::IUnknown>,
        fetched_count: *mut u32,
    ) -> windows_core::HRESULT {
        let fetched = match fetched_count.into_ref() {
            Ok(fetched) => fetched,
            Err(e) => return e.code(),
        };

        let elements = match range_elements.into_array_ref(count) {
            Ok(elements) => elements,
            Err(e) => return e.code(),
        };

        let mut items = Vec::with_capacity(count as usize);

        let code = self.0.next(count, fetched, &mut items);
        if code != S_OK {
            return code;
        }

        for (i, unk) in items.into_iter().enumerate() {
            elements[i] = Some(unk);
        }

        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        self.0.skip(count).ok()
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset().ok()
    }

    fn Clone(&self) -> windows_core::Result<IEnumUnknown> {
        Ok(IEnumUnknown::from(UnknownEnumerator(self.0.clone())))
    }
}

impl IEnumConnectionPoints_Impl for ConnectionPointsEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_connection_points: *mut Option<IConnectionPoint>,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let fetched = match count_fetched.into_ref() {
            Ok(fetched) => fetched,
            Err(e) => return e.code(),
        };

        let elements = match range_connection_points.into_array_ref(count) {
            Ok(elements) => elements,
            Err(e) => return e.code(),
        };

        let mut items = Vec::with_capacity(count as usize);

        let code = self.0.next(count, fetched, &mut items);
        if code != S_OK {
            return code;
        }

        for (i, cp) in items.into_iter().enumerate() {
            elements[i] = Some(cp);
        }

        S_OK
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        self.0.skip(count).ok()
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset().ok()
    }

    fn Clone(&self) -> windows_core::Result<IEnumConnectionPoints> {
        Ok(IEnumConnectionPoints::from(ConnectionPointsEnumerator(
            self.0.clone(),
        )))
    }
}

impl IEnumConnections_Impl for ConnectionsEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        range_connect_data: *mut CONNECTDATA,
        count_fetched: *mut u32,
    ) -> windows_core::HRESULT {
        let fetched = match count_fetched.into_ref() {
            Ok(fetched) => fetched,
            Err(e) => return e.code(),
        };

        let elements = match range_connect_data.into_array_ref(count) {
            Ok(elements) => elements,
            Err(e) => return e.code(),
        };

        self.0.next(count, fetched, elements)
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        self.0.skip(count).ok()
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset().ok()
    }

    fn Clone(&self) -> windows_core::Result<IEnumConnections> {
        Ok(IEnumConnections::from(ConnectionsEnumerator(
            self.0.clone(),
        )))
    }
}

impl opc_da_bindings::IEnumOPCItemAttributes_Impl for ItemAttributesEnumerator_Impl {
    fn Next(
        &self,
        count: u32,
        items: *mut *mut opc_da_bindings::tagOPCITEMATTRIBUTES,
        fetched_count: *mut u32,
    ) -> windows_core::Result<()> {
        let fetched = fetched_count.into_ref()?;
        let elements = items.into_com_array_ref(count)?;

        safe_call! {
            self.0.next(count, fetched, elements).ok(),
            items
        }
    }

    fn Skip(&self, count: u32) -> windows_core::Result<()> {
        self.0.skip(count).ok()
    }

    fn Reset(&self) -> windows_core::Result<()> {
        self.0.reset().ok()
    }

    fn Clone(&self) -> windows_core::Result<opc_da_bindings::IEnumOPCItemAttributes> {
        Ok(opc_da_bindings::IEnumOPCItemAttributes::from(
            ItemAttributesEnumerator(self.0.clone()),
        ))
    }
}
