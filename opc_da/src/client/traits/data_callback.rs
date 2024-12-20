use crate::client::RemoteArray;

#[windows::core::implement(
    // implicit implement IUnknown
    opc_da_bindings::IOPCDataCallback,
)]
pub struct DataCallback<'a, T>(pub &'a T)
where
    T: DataCallbackTrait + 'a;

impl<'a, T> std::ops::Deref for DataCallback<'a, T>
where
    T: DataCallbackTrait + 'a,
{
    type Target = T;

    fn deref(&self) -> &Self::Target {
        self.0
    }
}

pub enum DataCallbackEvent {
    DataChange {
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        values: RemoteArray<windows::core::VARIANT>,
        qualities: RemoteArray<u16>,
        timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
        errors: RemoteArray<windows_core::HRESULT>,
    },
    ReadComplete {
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        values: RemoteArray<windows::core::VARIANT>,
        qualities: RemoteArray<u16>,
        timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
        errors: RemoteArray<windows_core::HRESULT>,
    },
    WriteComplete {
        transaction_id: u32,
        group_handle: u32,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        errors: RemoteArray<windows_core::HRESULT>,
    },
    CancelComplete {
        transaction_id: u32,
        group_handle: u32,
    },
}

pub trait DataCallbackTrait {
    #[allow(clippy::too_many_arguments)]
    fn on_data_change(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        values: RemoteArray<windows::core::VARIANT>,
        qualities: RemoteArray<u16>,
        timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
        errors: RemoteArray<windows_core::HRESULT>,
    ) -> windows_core::Result<()>;

    #[allow(clippy::too_many_arguments)]
    fn on_read_complete(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        values: RemoteArray<windows::core::VARIANT>,
        qualities: RemoteArray<u16>,
        timestamps: RemoteArray<windows::Win32::Foundation::FILETIME>,
        errors: RemoteArray<windows_core::HRESULT>,
    ) -> windows_core::Result<()>;

    fn on_write_complete(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_error: windows_core::HRESULT,
        client_items: RemoteArray<u32>,
        errors: RemoteArray<windows_core::HRESULT>,
    ) -> windows_core::Result<()>;

    fn on_cancel_complete(
        &self,
        transaction_id: u32,
        group_handle: u32,
    ) -> windows_core::Result<()>;
}

impl<'a, T: DataCallbackTrait + 'a> opc_da_bindings::IOPCDataCallback_Impl
    for DataCallback_Impl<'a, T>
{
    fn OnDataChange(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        count: u32,
        client_items: *const u32,
        values: *const windows_core::VARIANT,
        qualities: *const u16,
        timestamps: *const windows::Win32::Foundation::FILETIME,
        errors: *const windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let client_items = RemoteArray::from_ptr(client_items, count);
        let values = RemoteArray::from_ptr(values, count);
        let qualities = RemoteArray::from_ptr(qualities, count);
        let timestamps = RemoteArray::from_ptr(timestamps, count);
        let errors = RemoteArray::from_ptr(errors, count);

        self.on_data_change(
            transaction_id,
            group_handle,
            master_quality,
            master_error,
            client_items,
            values,
            qualities,
            timestamps,
            errors,
        )
    }

    fn OnReadComplete(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_quality: windows_core::HRESULT,
        master_error: windows_core::HRESULT,
        count: u32,
        client_items: *const u32,
        values: *const windows_core::VARIANT,
        qualities: *const u16,
        timestamps: *const windows::Win32::Foundation::FILETIME,
        errors: *const windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let client_items = RemoteArray::from_ptr(client_items, count);
        let values = RemoteArray::from_ptr(values, count);
        let qualities = RemoteArray::from_ptr(qualities, count);
        let timestamps = RemoteArray::from_ptr(timestamps, count);
        let errors = RemoteArray::from_ptr(errors, count);

        self.on_read_complete(
            transaction_id,
            group_handle,
            master_quality,
            master_error,
            client_items,
            values,
            qualities,
            timestamps,
            errors,
        )
    }

    fn OnWriteComplete(
        &self,
        transaction_id: u32,
        group_handle: u32,
        master_error: windows_core::HRESULT,
        count: u32,
        client_handles: *const u32,
        errors: *const windows_core::HRESULT,
    ) -> windows_core::Result<()> {
        let client_items = RemoteArray::from_ptr(client_handles, count);
        let errors = RemoteArray::from_ptr(errors, count);

        self.on_write_complete(
            transaction_id,
            group_handle,
            master_error,
            client_items,
            errors,
        )
    }

    fn OnCancelComplete(&self, transaction_id: u32, group_handle: u32) -> windows_core::Result<()> {
        self.on_cancel_complete(transaction_id, group_handle)
    }
}
