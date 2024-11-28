use std::{
    collections::BTreeMap,
    mem::ManuallyDrop,
    sync::{atomic::AtomicU32, Arc},
};

use tokio::sync::RwLock;
use windows::Win32::{
    Foundation::{BOOL, E_INVALIDARG, E_NOTIMPL, FILETIME, S_OK},
    System::Com::{
        IAdviseSink, IConnectionPoint, IConnectionPointContainer, IConnectionPointContainer_Impl,
        IDataObject, IDataObject_Impl, IEnumConnectionPoints, IEnumFORMATETC, IEnumSTATDATA,
        FORMATETC, STGMEDIUM,
    },
};
use windows_core::{implement, ComObjectInner, VARIANT};

use super::base::{Core, Quality};

use super::{
    bindings,
    enumeration::ItemAttributesEnumerator,
    item::Item,
    utils::{com_alloc_v, copy_to_com_string},
};

#[implement(
    // implicit implement IUnknown
    bindings::IOPCItemMgt,
    bindings::IOPCGroupStateMgt,
    bindings::IOPCGroupStateMgt2,
    bindings::IOPCPublicGroupStateMgt,
    bindings::IOPCSyncIO,
    bindings::IOPCSyncIO2,
    bindings::IOPCAsyncIO2,
    bindings::IOPCAsyncIO3,
    bindings::IOPCItemDeadbandMgt,
    bindings::IOPCItemSamplingMgt,
    IConnectionPointContainer,
    bindings::IOPCAsyncIO,
    IDataObject
)]
pub struct Group {
    core: Arc<Core>,

    pub state: RwLock<GroupState>,

    // for IOPCItemMgt
    items_name_map: RwLock<BTreeMap<String, Arc<RwLock<Item>>>>,
    items_server_handle_map: RwLock<BTreeMap<u32, Arc<RwLock<Item>>>>,
    next_item_server_handle: AtomicU32,
}

pub struct GroupState {
    pub name: String,
    pub active: bool,
    pub update_rate: u32,
    pub client_group_handle: u32,
    pub time_bias: Option<i32>,
    pub percent_deadband: Option<f32>,
    pub locale_id: u32,
    pub server_group_handle: u32,
    pub keep_alive_time: u32,
}

impl Group {
    pub fn new(
        core: Arc<Core>,
        name: String,
        active: bool,
        update_rate: u32,
        client_group_handle: u32,
        time_bias: Option<i32>,
        percent_deadband: Option<f32>,
        locale_id: u32,
        server_group_handle: u32,
    ) -> Self {
        Self {
            core,
            state: RwLock::new(GroupState {
                name,
                active,
                update_rate,
                client_group_handle,
                time_bias,
                percent_deadband,
                locale_id,
                server_group_handle,
                keep_alive_time: 0,
            }),
            items_server_handle_map: RwLock::new(BTreeMap::new()),
            items_name_map: RwLock::new(BTreeMap::new()),
            next_item_server_handle: AtomicU32::new(1),
        }
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl bindings::IOPCItemMgt_Impl for Group_Impl {
    fn AddItems(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut r = Vec::new();
        let mut e = Vec::new();
        let item_name_map = self.items_name_map.blocking_write();
        for i in 0..count as usize {
            let item = unsafe { *item_array.add(i) };
            let name = unsafe { item.szItemID.to_string() }?;

            let item = match item_name_map.get(&name) {
                Some(item) => Some(item.clone()),
                None => {
                    let node = tokio::runtime::Handle::current()
                        .block_on(self.core.get_node_from_path(&name));

                    match node {
                        Some(node) => {
                            let server_handle = self
                                .next_item_server_handle
                                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
                            let item = Item {
                                name: name.clone(),
                                server_handle,
                                client_handle: item.hClient,
                                node: node.clone(),
                            };

                            let item = Arc::new(RwLock::new(item));
                            self.items_name_map
                                .blocking_write()
                                .insert(name, item.clone());
                            self.items_server_handle_map
                                .blocking_write()
                                .insert(server_handle, item.clone());

                            Some(item)
                        }
                        None => {
                            let result = bindings::tagOPCITEMRESULT {
                                hServer: 0,
                                vtCanonicalDataType: 0,
                                dwAccessRights: 0,
                                dwBlobSize: 0,
                                pBlob: std::ptr::null_mut(),
                                wReserved: 0,
                            };
                            r.push(result);
                            e.push(windows::Win32::Foundation::E_INVALIDARG);

                            None
                        }
                    }
                }
            };

            match item {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let result = bindings::tagOPCITEMRESULT {
                        hServer: item.server_handle,
                        vtCanonicalDataType: node.value.blocking_read().variant.get_data_type(),
                        dwAccessRights: 0,
                        dwBlobSize: 0,
                        pBlob: std::ptr::null_mut(),
                        wReserved: 0,
                    };
                    r.push(result);
                    e.push(windows::Win32::Foundation::S_OK);
                }
                None => {}
            }
        }
        unsafe {
            *results = com_alloc_v(&r);
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn ValidateItems(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        _blob_update: windows::Win32::Foundation::BOOL,
        validation_results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut r = Vec::new();
        let mut e = Vec::new();
        let item_name_map = self.items_name_map.blocking_write();
        for i in 0..count as usize {
            let item = unsafe { *item_array.add(i) };
            let name = unsafe { item.szItemID.to_string() }?;

            let item = match item_name_map.get(&name) {
                Some(item) => Some(item.clone()),
                None => None,
            };

            match item {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let result = bindings::tagOPCITEMRESULT {
                        hServer: item.server_handle,
                        vtCanonicalDataType: node.value.blocking_read().variant.get_data_type(),
                        dwAccessRights: 0,
                        dwBlobSize: 0,
                        pBlob: std::ptr::null_mut(),
                        wReserved: 0,
                    };
                    r.push(result);
                    e.push(windows::Win32::Foundation::S_OK);
                }
                None => {
                    let result = bindings::tagOPCITEMRESULT {
                        hServer: 0,
                        vtCanonicalDataType: 0,
                        dwAccessRights: 0,
                        dwBlobSize: 0,
                        pBlob: std::ptr::null_mut(),
                        wReserved: 0,
                    };
                    r.push(result);
                    e.push(windows::Win32::Foundation::E_INVALIDARG);
                }
            }
        }
        unsafe {
            *validation_results = com_alloc_v(&r);
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn RemoveItems(
        &self,
        count: u32,
        handle_server: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let mut item_server_handle_map = self.items_server_handle_map.blocking_write();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.remove(&server_handle) {
                Some(item) => {
                    let item = item.blocking_read();
                    self.items_name_map.blocking_write().remove(&item.name);
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            };
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn SetActiveState(
        &self,
        count: u32,
        handle_server: *const u32,
        active: windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_write();
                    item.node.blocking_write().state.blocking_write().is_active = active.as_bool();
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            };
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn SetClientHandles(
        &self,
        count: u32,
        handle_server: *const u32,
        handle_client: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            let client_handle = unsafe { *handle_client.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let mut item = item.blocking_write();
                    item.client_handle = client_handle;
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            };
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn SetDatatypes(
        &self,
        count: u32,
        _handle_server: *const u32,
        _requested_data_types: *const u16,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        // not surported
        let e = vec![E_INVALIDARG; count as usize];
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn CreateEnumerator(
        &self,
        _reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
        Ok(ItemAttributesEnumerator::new(
            self.items_name_map
                .blocking_read()
                .iter()
                .map(|(_, item)| {
                    tokio::runtime::Handle::current().block_on(async move {
                        let item = item.read().await;
                        let node = item.node.read().await;
                        let state = node.state.read().await;
                        let access_right = node.access_right.read().await;
                        let value = node.value.read().await;
                        bindings::tagOPCITEMATTRIBUTES {
                            szAccessPath: copy_to_com_string(""),
                            szItemID: copy_to_com_string(&node.get_path().await),
                            bActive: BOOL::from(state.is_active),
                            hClient: item.client_handle,
                            hServer: item.server_handle,
                            dwAccessRights: access_right.to_u32(),
                            dwBlobSize: 0,
                            pBlob: std::ptr::null_mut(),
                            vtRequestedDataType: value.variant.get_data_type(),
                            vtCanonicalDataType: value.variant.get_data_type(),
                            dwEUType: bindings::OPC_NOENUM,
                            vEUInfo: ManuallyDrop::new(VARIANT::new()),
                        }
                    })
                })
                .collect(),
        )
        .into_object()
        .into_interface())
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl bindings::IOPCGroupStateMgt_Impl for Group_Impl {
    fn GetState(
        &self,
        update_rate: *mut u32,
        active: *mut windows::Win32::Foundation::BOOL,
        name: *mut windows_core::PWSTR,
        timebias: *mut i32,
        percent_deadband: *mut f32,
        locale_id: *mut u32,
        handle_client_group: *mut u32,
        handle_server_group: *mut u32,
    ) -> windows_core::Result<()> {
        let state = self.state.blocking_read();
        unsafe {
            *update_rate = state.update_rate;
            *active = state.active.into();
            *name = copy_to_com_string(&state.name);
            *timebias = state.time_bias.unwrap_or(0);
            *percent_deadband = state.percent_deadband.unwrap_or(0.0);
            *locale_id = state.locale_id;
            *handle_client_group = state.client_group_handle;
            *handle_server_group = state.server_group_handle;
        }
        Ok(())
    }

    fn SetState(
        &self,
        _requested_update_rate: *const u32,
        revised_update_rate: *mut u32,
        _active: *const windows::Win32::Foundation::BOOL,
        _timebias: *const i32,
        _percent_deadband: *const f32,
        _locale_id: *const u32,
        _handle_client_group: *const u32,
    ) -> windows_core::Result<()> {
        let state = self.state.blocking_read();
        // TODO: ignore others now
        unsafe {
            *revised_update_rate = state.update_rate;
        }
        Ok(())
    }

    fn SetName(&self, _name: &windows_core::PCWSTR) -> windows_core::Result<()> {
        // TODO: ignore now
        Ok(())
    }

    fn CloneGroup(
        &self,
        _name: &windows_core::PCWSTR,
        _reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
        Err(E_INVALIDARG.into())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl bindings::IOPCGroupStateMgt2_Impl for Group_Impl {
    fn SetKeepAlive(&self, keep_alive_time: u32) -> windows_core::Result<u32> {
        let mut state = self.state.blocking_write();
        state.keep_alive_time = keep_alive_time;
        Ok(keep_alive_time)
    }

    fn GetKeepAlive(&self) -> windows_core::Result<u32> {
        let state = self.state.blocking_read();
        Ok(state.keep_alive_time)
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl bindings::IOPCPublicGroupStateMgt_Impl for Group_Impl {
    fn GetState(&self) -> windows_core::Result<windows::Win32::Foundation::BOOL> {
        Err(E_NOTIMPL.into())
    }

    fn MoveToPublic(&self) -> windows_core::Result<()> {
        Err(E_NOTIMPL.into())
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl bindings::IOPCSyncIO_Impl for Group_Impl {
    fn Read(
        &self,
        _source: bindings::tagOPCDATASOURCE,
        count: u32,
        handle_server: *const u32,
        item_values: *mut *mut bindings::tagOPCITEMSTATE,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut v = Vec::new();
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let value = node.value.blocking_read();
                    let state = bindings::tagOPCITEMSTATE {
                        hClient: item.client_handle,
                        ftTimeStamp: value
                            .timestamp
                            .clone()
                            .map_or_else(FILETIME::default, |t| t.into()),
                        wQuality: 0,
                        vDataValue: ManuallyDrop::new(value.variant.clone().into()),
                        wReserved: 0,
                    };
                    v.push(state);
                    e.push(S_OK);
                }
                None => {
                    let state = bindings::tagOPCITEMSTATE {
                        hClient: 0,
                        ftTimeStamp: FILETIME::default(),
                        wQuality: 0,
                        vDataValue: ManuallyDrop::new(VARIANT::new()),
                        wReserved: 0,
                    };
                    v.push(state);
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *item_values = com_alloc_v(&v);
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn Write(
        &self,
        count: u32,
        handle_server: *const u32,
        item_values: *const windows_core::VARIANT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            let value = unsafe { (*item_values.add(i)).clone() };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_write();
                    let node = item.node.blocking_write();
                    node.value.blocking_write().variant = value.into();
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl bindings::IOPCSyncIO2_Impl for Group_Impl {
    fn ReadMaxAge(
        &self,
        count: u32,
        handle_server: *const u32,
        _max_age: *const u32,
        values: *mut *mut windows_core::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut v = Vec::new();
        let mut q = Vec::new();
        let mut t = Vec::new();
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let value = node.value.blocking_read();
                    v.push(value.variant.clone().into());
                    q.push(value.quality.to_u16());
                    t.push(
                        value
                            .timestamp
                            .clone()
                            .map_or_else(FILETIME::default, |t| t.into()),
                    );
                    e.push(S_OK);
                }
                None => {
                    v.push(VARIANT::new());
                    q.push(0);
                    t.push(FILETIME::default());
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *values = com_alloc_v(&v);
            *qualities = com_alloc_v(&q);
            *timestamps = com_alloc_v(&t);
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn WriteVQT(
        &self,
        count: u32,
        handle_server: *const u32,
        item_vqt: *const bindings::tagOPCITEMVQT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            let vqt = unsafe { &*item_vqt.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_write();
                    let node = item.node.blocking_write();
                    let mut value = node.value.blocking_write();
                    value.variant = vqt.vDataValue.as_raw().clone().into();
                    value.quality = Quality(vqt.wQuality);
                    value.timestamp = Some(vqt.ftTimeStamp.clone().into());
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl bindings::IOPCAsyncIO2_Impl for Group_Impl {
    fn Read(
        &self,
        count: u32,
        handle_server: *const u32,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let value = node.value.blocking_read();
                    let state = bindings::tagOPCITEMSTATE {
                        hClient: item.client_handle,
                        ftTimeStamp: value
                            .timestamp
                            .clone()
                            .map_or_else(FILETIME::default, |t| t.into()),
                        wQuality: 0,
                        vDataValue: ManuallyDrop::new(value.variant.clone().into()),
                        wReserved: 0,
                    };
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn Write(
        &self,
        count: u32,
        handle_server: *const u32,
        item_values: *const windows_core::VARIANT,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            let value = unsafe { (*item_values.add(i)).clone() };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_write();
                    let node = item.node.blocking_write();
                    node.value.blocking_write().variant = value.into();
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn Refresh2(
        &self,
        _source: bindings::tagOPCDATASOURCE,
        _transaction_id: u32,
    ) -> windows_core::Result<u32> {
        Ok(S_OK.0 as u32)
    }

    fn Cancel2(&self, _cancel_id: u32) -> windows_core::Result<()> {
        Ok(())
    }

    fn SetEnable(&self, enable: windows::Win32::Foundation::BOOL) -> windows_core::Result<()> {
        let mut state = self.state.blocking_write();
        state.active = enable.as_bool();
        Ok(())
    }

    fn GetEnable(&self) -> windows_core::Result<windows::Win32::Foundation::BOOL> {
        let state = self.state.blocking_read();
        Ok(state.active.into())
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl bindings::IOPCAsyncIO3_Impl for Group_Impl {
    fn ReadMaxAge(
        &self,
        count: u32,
        handle_server: *const u32,
        _max_age: *const u32,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let mut e = Vec::new();
        let item_server_handle_map = self.items_server_handle_map.blocking_read();
        for i in 0..count as usize {
            let server_handle = unsafe { *handle_server.add(i) };
            match item_server_handle_map.get(&server_handle) {
                Some(item) => {
                    let item = item.blocking_read();
                    let node = item.node.blocking_read();
                    let value = node.value.blocking_read();
                    let state = bindings::tagOPCITEMSTATE {
                        hClient: item.client_handle,
                        ftTimeStamp: value
                            .timestamp
                            .clone()
                            .map_or_else(FILETIME::default, |t| t.into()),
                        wQuality: 0,
                        vDataValue: ManuallyDrop::new(value.variant.clone().into()),
                        wReserved: 0,
                    };
                    e.push(S_OK);
                }
                None => {
                    e.push(E_INVALIDARG);
                }
            }
        }
        unsafe {
            *errors = com_alloc_v(&e);
        }
        Ok(())
    }

    fn WriteVQT(
        &self,
        count: u32,
        handle_server: *const u32,
        item_vqt: *const bindings::tagOPCITEMVQT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn RefreshMaxAge(&self, max_age: u32, transaction_id: u32) -> windows_core::Result<u32> {
        todo!()
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl bindings::IOPCItemDeadbandMgt_Impl for Group_Impl {
    fn SetItemDeadband(
        &self,
        count: u32,
        handle_server: *const u32,
        percent_deadband: *const f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn GetItemDeadband(
        &self,
        count: u32,
        handle_server: *const u32,
        percent_deadband: *mut *mut f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn ClearItemDeadband(
        &self,
        count: u32,
        handle_server: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 optional
impl bindings::IOPCItemSamplingMgt_Impl for Group_Impl {
    fn SetItemSamplingRate(
        &self,
        count: u32,
        handle_server: *const u32,
        requested_sampling_rate: *const u32,
        revised_sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn GetItemSamplingRate(
        &self,
        count: u32,
        handle_server: *const u32,
        sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn ClearItemSamplingRate(
        &self,
        count: u32,
        handle_server: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn SetItemBufferEnable(
        &self,
        count: u32,
        handle_server: *const u32,
        penable: *const windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn GetItemBufferEnable(
        &self,
        count: u32,
        handle_server: *const u32,
        enable: *mut *mut windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl IConnectionPointContainer_Impl for Group_Impl {
    fn EnumConnectionPoints(&self) -> windows_core::Result<IEnumConnectionPoints> {
        todo!()
    }

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<IConnectionPoint> {
        todo!()
    }
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl bindings::IOPCAsyncIO_Impl for Group_Impl {
    fn Read(
        &self,
        connection: u32,
        source: bindings::tagOPCDATASOURCE,
        count: u32,
        handle_server: *const u32,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn Write(
        &self,
        connection: u32,
        count: u32,
        handle_server: *const u32,
        item_values: *const windows_core::VARIANT,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn Refresh(
        &self,
        connection: u32,
        source: bindings::tagOPCDATASOURCE,
    ) -> windows_core::Result<u32> {
        todo!()
    }

    fn Cancel(&self, transaction_id: u32) -> windows_core::Result<()> {
        todo!()
    }
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl IDataObject_Impl for Group_Impl {
    fn GetData(&self, format_etc_in: *const FORMATETC) -> windows_core::Result<STGMEDIUM> {
        todo!()
    }

    fn GetDataHere(
        &self,
        format_etc_in: *const FORMATETC,
        medium: *mut STGMEDIUM,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn QueryGetData(&self, format_etc_in: *const FORMATETC) -> windows_core::HRESULT {
        todo!()
    }

    fn GetCanonicalFormatEtc(
        &self,
        format_etc_in: *const FORMATETC,
        format_etc_inout: *mut FORMATETC,
    ) -> windows_core::HRESULT {
        todo!()
    }

    fn SetData(
        &self,
        format_etc_in: *const FORMATETC,
        medium: *const STGMEDIUM,
        release: windows::Win32::Foundation::BOOL,
    ) -> windows_core::Result<()> {
        todo!()
    }

    fn EnumFormatEtc(&self, direction: u32) -> windows_core::Result<IEnumFORMATETC> {
        todo!()
    }

    fn DAdvise(
        &self,
        format_etc_in: *const FORMATETC,
        adv: u32,
        sink: Option<&IAdviseSink>,
    ) -> windows_core::Result<u32> {
        todo!()
    }

    fn DUnadvise(&self, connection: u32) -> windows_core::Result<()> {
        todo!()
    }

    fn EnumDAdvise(&self) -> windows_core::Result<IEnumSTATDATA> {
        todo!()
    }
}
