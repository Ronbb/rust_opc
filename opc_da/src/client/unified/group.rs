use std::collections::{BTreeMap, HashMap};

use windows_core::{ComObjectInner as _, IUnknown, Interface};

use crate::{
    client::{
        v1, v2, v3, AsyncIo2Trait, AsyncIo3Trait, ConnectionPointContainerTrait, DataCallback,
        DataCallbackTrait, ItemMgtTrait, SyncIo2Trait, SyncIoTrait,
    },
    def::{
        CancelCompleteEvent, DataChangeEvent, DataSourceTarget, ItemDef, ItemResult, ItemState,
        ItemValue, ReadCompleteEvent, WriteCompleteEvent,
    },
    utils::{IntoBridge as _, TryToLocal as _, TryToNative as _},
};

pub struct Group {
    inner: GroupInner,
    items: HashMap<String, Item>,
    next_transaction_id: std::sync::atomic::AtomicU32,
    initialized: bool,
    data_callback_cookie: Option<u32>,
    data_change_broadcaster: tokio::sync::broadcast::Sender<DataChangeEvent>,
    data_change_awaiters:
        std::sync::Mutex<BTreeMap<u32, tokio::sync::oneshot::Sender<DataChangeEvent>>>,
    read_complete_awaiters:
        std::sync::Mutex<BTreeMap<u32, tokio::sync::oneshot::Sender<ReadCompleteEvent>>>,
    write_complete_awaiters:
        std::sync::Mutex<BTreeMap<u32, tokio::sync::oneshot::Sender<WriteCompleteEvent>>>,
    cancel_complete_awaiters:
        std::sync::Mutex<BTreeMap<u32, tokio::sync::oneshot::Sender<CancelCompleteEvent>>>,
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
    fn new(inner: GroupInner) -> Self {
        let data_change_broadcaster = tokio::sync::broadcast::Sender::new(32);

        Self {
            inner,
            items: HashMap::new(),
            next_transaction_id: std::sync::atomic::AtomicU32::new(1),
            initialized: false,
            data_callback_cookie: None,
            data_change_broadcaster,
            data_change_awaiters: std::sync::Mutex::new(BTreeMap::new()),
            read_complete_awaiters: std::sync::Mutex::new(BTreeMap::new()),
            write_complete_awaiters: std::sync::Mutex::new(BTreeMap::new()),
            cancel_complete_awaiters: std::sync::Mutex::new(BTreeMap::new()),
        }
    }

    pub fn initialize(&mut self) -> windows::core::Result<()> {
        if self.initialized {
            return Ok(());
        }

        let connection_point = match &self.inner {
            GroupInner::V1(_) => return Ok(()),
            GroupInner::V2(group) => group.data_callback_connection_point()?,
            GroupInner::V3(group) => group.data_callback_connection_point()?,
        };

        if self.data_callback_cookie.is_none() {
            let callback = DataCallback(self);
            self.data_callback_cookie = Some(unsafe {
                connection_point.Advise(
                    &callback
                        .into_object()
                        .into_interface::<opc_da_bindings::IOPCDataCallback>()
                        .cast::<IUnknown>()?,
                )
            }?);
        }

        self.initialized = true;

        Ok(())
    }

    pub fn data_change_receiver(&self) -> tokio::sync::broadcast::Receiver<DataChangeEvent> {
        self.data_change_broadcaster.subscribe()
    }
}

impl DataCallbackTrait for Group {
    fn on_data_change(&self, event: DataChangeEvent) -> windows_core::Result<()> {
        self.data_change_broadcaster
            .send(event.clone())
            .map_err(|_| {
                windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "data change event receiver dropped",
                )
            })?;

        let mut awaiters = self.data_change_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        let awaiter = match awaiters.remove(&event.transaction_id) {
            Some(awaiter) => awaiter,
            None => {
                return Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "no awaiter found",
                ))
            }
        };

        awaiter.send(event).map_err(|_| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_FAIL,
                "data change event awaiter dropped",
            )
        })?;

        Ok(())
    }

    fn on_read_complete(&self, event: ReadCompleteEvent) -> windows_core::Result<()> {
        let mut awaiters = self.read_complete_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        let awaiter = match awaiters.remove(&event.transaction_id) {
            Some(awaiter) => awaiter,
            None => {
                return Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "no awaiter found",
                ))
            }
        };

        awaiter.send(event).map_err(|_| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_FAIL,
                "read complete event awaiter dropped",
            )
        })?;

        Ok(())
    }

    fn on_write_complete(&self, event: WriteCompleteEvent) -> windows_core::Result<()> {
        let mut awaiters = self.write_complete_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        let awaiter = match awaiters.remove(&event.transaction_id) {
            Some(awaiter) => awaiter,
            None => {
                return Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "no awaiter found",
                ))
            }
        };

        awaiter.send(event).map_err(|_| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_FAIL,
                "write complete event awaiter dropped",
            )
        })?;

        Ok(())
    }

    fn on_cancel_complete(&self, event: CancelCompleteEvent) -> windows_core::Result<()> {
        let mut awaiters = self.cancel_complete_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        let awaiter = match awaiters.remove(&event.transaction_id) {
            Some(awaiter) => awaiter,
            None => {
                return Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "no awaiter found",
                ))
            }
        };

        awaiter.send(event).map_err(|_| {
            windows_core::Error::new(
                windows::Win32::Foundation::E_FAIL,
                "cancel complete event awaiter dropped",
            )
        })?;

        Ok(())
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

    fn read_items_sync1<T: SyncIoTrait>(
        &self,
        sync_io1: &T,
        data_source: DataSourceTarget,
        server_handles: &[u32],
    ) -> windows::core::Result<Vec<windows::core::Result<ItemValue>>> {
        let results: Vec<windows::core::Result<ItemState>> = sync_io1
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

    pub fn read_items_sync<S>(
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
                    .get(name.as_ref())
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
            GroupInner::V1(group) => self.read_items_sync1(group, data_source, &server_handles),
            GroupInner::V2(group) => self.read_items_sync1(group, data_source, &server_handles),
            GroupInner::V3(group) => self.read_items_sync2(
                group,
                &server_handles,
                &vec![data_source.max_age(); server_handles.len()],
            ),
        }
    }

    fn read_items_async2<T: AsyncIo2Trait>(
        &self,
        async_io2: &T,
        server_handles: &[u32],
    ) -> windows::core::Result<(
        DataCallbackFuture<ReadCompleteEvent>,
        Vec<windows::core::Result<()>>,
    )> {
        let transaction_id = self
            .next_transaction_id
            .fetch_add(1, std::sync::atomic::Ordering::SeqCst);

        let (sender, receive) = tokio::sync::oneshot::channel();

        let mut awaiters = self.read_complete_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        awaiters.insert(transaction_id, sender);

        let (cancel_id, results) = async_io2.read(server_handles, transaction_id)?;

        Ok((
            DataCallbackFuture {
                receiver: Box::pin(receive),
                transaction_id,
                cancel_id,
            },
            results.try_to_local()?,
        ))
    }

    fn read_items_async3<T: AsyncIo3Trait>(
        &self,
        async_io3: &T,
        server_handles: &[u32],
        max_ages: &[u32],
    ) -> windows::core::Result<(
        DataCallbackFuture<ReadCompleteEvent>,
        Vec<windows::core::Result<()>>,
    )> {
        let transaction_id = self
            .next_transaction_id
            .fetch_add(1, std::sync::atomic::Ordering::SeqCst);

        let (sender, receive) = tokio::sync::oneshot::channel();

        let mut awaiters = self.read_complete_awaiters.lock().map_err(|_| {
            windows_core::Error::new(windows::Win32::Foundation::E_FAIL, "lock poisoned")
        })?;

        awaiters.insert(transaction_id, sender);

        let (cancel_id, results) =
            async_io3.read_max_age(server_handles, max_ages, transaction_id)?;

        Ok((
            DataCallbackFuture {
                receiver: Box::pin(receive),
                transaction_id,
                cancel_id,
            },
            results.try_to_local()?,
        ))
    }

    pub fn read_items_async<S: AsRef<str>>(
        &self,
        items_names: &[S],
        data_source: DataSourceTarget,
    ) -> windows::core::Result<(
        DataCallbackFuture<ReadCompleteEvent>,
        Vec<windows::core::Result<()>>,
    )> {
        let server_handles: Vec<u32> = items_names
            .iter()
            .map(|name| {
                self.items
                    .get(name.as_ref())
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
            GroupInner::V1(_) => Err(windows_core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "read_items_async not implemented for v1",
            )),
            GroupInner::V2(group) => self.read_items_async2(group, &server_handles),
            GroupInner::V3(group) => self.read_items_async3(
                group,
                &server_handles,
                &vec![data_source.max_age(); server_handles.len()],
            ),
        }
    }
}

impl From<v1::Group> for Group {
    fn from(group: v1::Group) -> Self {
        Self::new(GroupInner::V1(group))
    }
}

impl From<v2::Group> for Group {
    fn from(group: v2::Group) -> Self {
        Self::new(GroupInner::V2(group))
    }
}

impl From<v3::Group> for Group {
    fn from(group: v3::Group) -> Self {
        Self::new(GroupInner::V3(group))
    }
}

pub struct DataCallbackFuture<T> {
    receiver: std::pin::Pin<Box<tokio::sync::oneshot::Receiver<T>>>,
    transaction_id: u32,
    cancel_id: u32,
}

impl<T> DataCallbackFuture<T> {
    pub fn cancel_id(&self) -> u32 {
        self.cancel_id
    }

    pub fn transaction_id(&self) -> u32 {
        self.transaction_id
    }
}

impl<T> std::future::Future for DataCallbackFuture<T> {
    type Output = windows::core::Result<T>;

    fn poll(
        mut self: std::pin::Pin<&mut Self>,
        cx: &mut std::task::Context<'_>,
    ) -> std::task::Poll<Self::Output> {
        match self.receiver.as_mut().poll(cx) {
            std::task::Poll::Ready(Ok(event)) => std::task::Poll::Ready(Ok(event)),
            std::task::Poll::Ready(Err(_)) => {
                std::task::Poll::Ready(Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_FAIL,
                    "data change event receiver dropped",
                )))
            }
            std::task::Poll::Pending => std::task::Poll::Pending,
        }
    }
}
