use std::ops::Deref;

use windows::Win32::System::Com::{
    IAdviseSink, IConnectionPoint, IConnectionPointContainer, IConnectionPointContainer_Impl,
    IDataObject, IDataObject_Impl, IEnumConnectionPoints, IEnumFORMATETC, IEnumSTATDATA, FORMATETC,
    STGMEDIUM,
};
use windows_core::implement;

use crate::traits::GroupTrait;

use super::bindings;

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
pub struct Group<T>(pub T)
where
    T: GroupTrait + 'static;

impl<T: GroupTrait + 'static> Deref for Group<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCItemMgt_Impl for Group_Impl<T> {
    fn AddItems(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn ValidateItems(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        _blob_update: windows::Win32::Foundation::BOOL,
        validation_results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn RemoveItems(
        &self,
        count: u32,
        item_server_handle: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn SetActiveState(
        &self,
        count: u32,
        item_server_handle: *const u32,
        active: windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn SetClientHandles(
        &self,
        count: u32,
        item_server_handle: *const u32,
        handle_client: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn SetDatatypes(
        &self,
        count: u32,
        _item_server_handle: *const u32,
        _requested_data_types: *const u16,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn CreateEnumerator(
        &self,
        _reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCGroupStateMgt_Impl for Group_Impl<T> {
    fn GetState(
        &self,
        update_rate: *mut u32,
        active: *mut windows::Win32::Foundation::BOOL,
        name: *mut windows_core::PWSTR,
        timebias: *mut i32,
        percent_deadband: *mut f32,
        locale_id: *mut u32,
        group_client_handle: *mut u32,
        item_server_handle_group: *mut u32,
    ) -> windows_core::Result<()> {
    }

    fn SetState(
        &self,
        _requested_update_rate: *const u32,
        revised_update_rate: *mut u32,
        _active: *const windows::Win32::Foundation::BOOL,
        _timebias: *const i32,
        _percent_deadband: *const f32,
        _locale_id: *const u32,
        _group_client_handle: *const u32,
    ) -> windows_core::Result<()> {
    }

    fn SetName(&self, _name: &windows_core::PCWSTR) -> windows_core::Result<()> {}

    fn CloneGroup(
        &self,
        _name: &windows_core::PCWSTR,
        _reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown> {
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCGroupStateMgt2_Impl for Group_Impl<T> {
    fn SetKeepAlive(&self, keep_alive_time: u32) -> windows_core::Result<u32> {}

    fn GetKeepAlive(&self) -> windows_core::Result<u32> {}
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> bindings::IOPCPublicGroupStateMgt_Impl for Group_Impl<T> {
    fn GetState(&self) -> windows_core::Result<windows::Win32::Foundation::BOOL> {}

    fn MoveToPublic(&self) -> windows_core::Result<()> {}
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCSyncIO_Impl for Group_Impl<T> {
    fn Read(
        &self,
        _source: bindings::tagOPCDATASOURCE,
        count: u32,
        item_server_handle: *const u32,
        item_values: *mut *mut bindings::tagOPCITEMSTATE,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn Write(
        &self,
        count: u32,
        item_server_handle: *const u32,
        item_values: *const windows_core::VARIANT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCSyncIO2_Impl for Group_Impl<T> {
    fn ReadMaxAge(
        &self,
        count: u32,
        item_server_handle: *const u32,
        _max_age: *const u32,
        values: *mut *mut windows_core::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_server_handle: *const u32,
        item_vqt: *const bindings::tagOPCITEMVQT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCAsyncIO2_Impl for Group_Impl<T> {
    fn Read(
        &self,
        count: u32,
        item_server_handle: *const u32,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn Write(
        &self,
        count: u32,
        item_server_handle: *const u32,
        item_values: *const windows_core::VARIANT,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn Refresh2(
        &self,
        _source: bindings::tagOPCDATASOURCE,
        _transaction_id: u32,
    ) -> windows_core::Result<u32> {
    }

    fn Cancel2(&self, _cancel_id: u32) -> windows_core::Result<()> {}

    fn SetEnable(&self, enable: windows::Win32::Foundation::BOOL) -> windows_core::Result<()> {}

    fn GetEnable(&self) -> windows_core::Result<windows::Win32::Foundation::BOOL> {}
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCAsyncIO3_Impl for Group_Impl<T> {
    fn ReadMaxAge(
        &self,
        count: u32,
        item_server_handle: *const u32,
        _max_age: *const u32,
        _transaction_id: u32,
        _cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_server_handle: *const u32,
        item_vqt: *const bindings::tagOPCITEMVQT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn RefreshMaxAge(&self, max_age: u32, transaction_id: u32) -> windows_core::Result<u32> {}
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> bindings::IOPCItemDeadbandMgt_Impl for Group_Impl<T> {
    fn SetItemDeadband(
        &self,
        count: u32,
        item_server_handle: *const u32,
        percent_deadband: *const f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn GetItemDeadband(
        &self,
        count: u32,
        item_server_handle: *const u32,
        percent_deadband: *mut *mut f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn ClearItemDeadband(
        &self,
        count: u32,
        item_server_handle: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 optional
impl<T: GroupTrait + 'static> bindings::IOPCItemSamplingMgt_Impl for Group_Impl<T> {
    fn SetItemSamplingRate(
        &self,
        count: u32,
        item_server_handle: *const u32,
        requested_sampling_rate: *const u32,
        revised_sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn GetItemSamplingRate(
        &self,
        count: u32,
        item_server_handle: *const u32,
        sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn ClearItemSamplingRate(
        &self,
        count: u32,
        item_server_handle: *const u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn SetItemBufferEnable(
        &self,
        count: u32,
        item_server_handle: *const u32,
        penable: *const windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn GetItemBufferEnable(
        &self,
        count: u32,
        item_server_handle: *const u32,
        enable: *mut *mut windows::Win32::Foundation::BOOL,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> IConnectionPointContainer_Impl for Group_Impl<T> {
    fn EnumConnectionPoints(&self) -> windows_core::Result<IEnumConnectionPoints> {}

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<IConnectionPoint> {
    }
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> bindings::IOPCAsyncIO_Impl for Group_Impl<T> {
    fn Read(
        &self,
        connection: u32,
        source: bindings::tagOPCDATASOURCE,
        count: u32,
        item_server_handle: *const u32,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn Write(
        &self,
        connection: u32,
        count: u32,
        item_server_handle: *const u32,
        item_values: *const windows_core::VARIANT,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()> {
    }

    fn Refresh(
        &self,
        connection: u32,
        source: bindings::tagOPCDATASOURCE,
    ) -> windows_core::Result<u32> {
    }

    fn Cancel(&self, transaction_id: u32) -> windows_core::Result<()> {}
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> IDataObject_Impl for Group_Impl<T> {
    fn GetData(&self, format_etc_in: *const FORMATETC) -> windows_core::Result<STGMEDIUM> {}

    fn GetDataHere(
        &self,
        format_etc_in: *const FORMATETC,
        medium: *mut STGMEDIUM,
    ) -> windows_core::Result<()> {
    }

    fn QueryGetData(&self, format_etc_in: *const FORMATETC) -> windows_core::HRESULT {}

    fn GetCanonicalFormatEtc(
        &self,
        format_etc_in: *const FORMATETC,
        format_etc_inout: *mut FORMATETC,
    ) -> windows_core::HRESULT {
    }

    fn SetData(
        &self,
        format_etc_in: *const FORMATETC,
        medium: *const STGMEDIUM,
        release: windows::Win32::Foundation::BOOL,
    ) -> windows_core::Result<()> {
    }

    fn EnumFormatEtc(&self, direction: u32) -> windows_core::Result<IEnumFORMATETC> {}

    fn DAdvise(
        &self,
        format_etc_in: *const FORMATETC,
        adv: u32,
        sink: Option<&IAdviseSink>,
    ) -> windows_core::Result<u32> {
    }

    fn DUnadvise(&self, connection: u32) -> windows_core::Result<()> {}

    fn EnumDAdvise(&self) -> windows_core::Result<IEnumSTATDATA> {}
}
