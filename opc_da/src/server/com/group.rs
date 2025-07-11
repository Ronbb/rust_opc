use crate::{
    safe_call,
    server::{
        com::memory::{FreeRaw as _, IntoRef as _},
        traits::GroupTrait,
    },
};

use super::memory::IntoComArrayRef;

#[windows::core::implement(
    // implicit implement IUnknown
    opc_da_bindings::IOPCItemMgt,
    opc_da_bindings::IOPCGroupStateMgt,
    opc_da_bindings::IOPCGroupStateMgt2,
    opc_da_bindings::IOPCPublicGroupStateMgt,
    opc_da_bindings::IOPCSyncIO,
    opc_da_bindings::IOPCSyncIO2,
    opc_da_bindings::IOPCAsyncIO2,
    opc_da_bindings::IOPCAsyncIO3,
    opc_da_bindings::IOPCItemDeadbandMgt,
    opc_da_bindings::IOPCItemSamplingMgt,
    windows::Win32::System::Com::IConnectionPointContainer,
    opc_da_bindings::IOPCAsyncIO,
    windows::Win32::System::Com::IDataObject
)]
pub struct Group<T>(pub T)
where
    T: GroupTrait + 'static;

impl<T: GroupTrait + 'static> core::ops::Deref for Group<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCItemMgt_Impl for Group_Impl<T> {
    fn AddItems(
        &self,
        count: u32,
        items: *const opc_da_bindings::tagOPCITEMDEF,
        results: *mut *mut opc_da_bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.add_items(
                items.into_com_array_ref(count)?,
                results.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            results,
            errors
        }
    }

    fn ValidateItems(
        &self,
        count: u32,
        items: *const opc_da_bindings::tagOPCITEMDEF,
        blob_update: windows_core::BOOL,
        validation_results: *mut *mut opc_da_bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.validate_items(
                items.into_com_array_ref(count)?,
                blob_update,
                validation_results.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            validation_results,
            errors
        }
    }

    fn RemoveItems(
        &self,
        count: u32,
        item_server_handles: *const u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.remove_items(
                item_server_handles.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn SetActiveState(
        &self,
        count: u32,
        item_server_handles: *const u32,
        active: windows_core::BOOL,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_active_state(
                item_server_handles.into_com_array_ref(count)?,
                active,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn SetClientHandles(
        &self,
        count: u32,
        item_server_handles: *const u32,
        handle_client: *const u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_client_handles(
                item_server_handles.into_com_array_ref(count)?,
                handle_client.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn SetDatatypes(
        &self,
        count: u32,
        item_server_handles: *const u32,
        requested_data_types: *const u16,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_data_types(
                item_server_handles.into_com_array_ref(count)?,
                requested_data_types.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn CreateEnumerator(
        &self,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        safe_call! {
            self.create_enumerator(reference_interface_id.into_ref()?),
        }
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCGroupStateMgt_Impl for Group_Impl<T> {
    fn GetState(
        &self,
        update_rate: *mut u32,
        active: *mut windows_core::BOOL,
        name: *mut windows::core::PWSTR,
        time_bias: *mut i32,
        percent_deadband: *mut f32,
        locale_id: *mut u32,
        group_client_handle: *mut u32,
        item_server_handle_group: *mut u32,
    ) -> windows::core::Result<()> {
        self.get_state(
            update_rate.into_ref()?,
            active.into_ref()?,
            name.into_ref()?,
            time_bias.into_ref()?,
            percent_deadband.into_ref()?,
            locale_id.into_ref()?,
            group_client_handle.into_ref()?,
            item_server_handle_group.into_ref()?,
        )
    }

    fn SetState(
        &self,
        requested_update_rate: *const u32,
        revised_update_rate: *mut u32,
        active: *const windows_core::BOOL,
        time_bias: *const i32,
        percent_deadband: *const f32,
        locale_id: *const u32,
        group_client_handle: *const u32,
    ) -> windows::core::Result<()> {
        self.set_state(
            requested_update_rate.into_ref()?,
            revised_update_rate.into_ref()?,
            active.into_ref()?,
            time_bias.into_ref()?,
            percent_deadband.into_ref()?,
            locale_id.into_ref()?,
            group_client_handle.into_ref()?,
        )
    }

    fn SetName(&self, name: &windows::core::PCWSTR) -> windows::core::Result<()> {
        self.set_name(name)
    }

    fn CloneGroup(
        &self,
        name: &windows::core::PCWSTR,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown> {
        self.clone_group(name, reference_interface_id.into_ref()?)
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCGroupStateMgt2_Impl for Group_Impl<T> {
    fn SetKeepAlive(&self, keep_alive_time: u32) -> windows::core::Result<u32> {
        self.set_keep_alive(keep_alive_time)
    }

    fn GetKeepAlive(&self) -> windows::core::Result<u32> {
        self.get_keep_alive()
    }
}

// 1.0 optional
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCPublicGroupStateMgt_Impl for Group_Impl<T> {
    fn GetState(&self) -> windows::core::Result<windows_core::BOOL> {
        self.get_public_group_state()
    }

    fn MoveToPublic(&self) -> windows::core::Result<()> {
        self.move_to_public()
    }
}

// 1.0 required
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCSyncIO_Impl for Group_Impl<T> {
    fn Read(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        count: u32,
        item_server_handles: *const u32,
        item_values: *mut *mut opc_da_bindings::tagOPCITEMSTATE,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.read(
                source,
                item_server_handles.into_com_array_ref(count)?,
                item_values.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            item_values,
            errors
        }
    }

    fn Write(
        &self,
        count: u32,
        item_server_handles: *const u32,
        item_values: *const windows::Win32::System::Variant::VARIANT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.write(
                item_server_handles.into_com_array_ref(count)?,
                item_values.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCSyncIO2_Impl for Group_Impl<T> {
    fn ReadMaxAge(
        &self,
        count: u32,
        item_server_handles: *const u32,
        max_age: *const u32,
        values: *mut *mut windows::Win32::System::Variant::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.read_max_age(
                item_server_handles.into_com_array_ref(count)?,
                max_age.into_com_array_ref(count)?,
                values.into_com_array_ref(count)?,
                qualities.into_com_array_ref(count)?,
                timestamps.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            values,
            qualities,
            timestamps,
            errors
        }
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_server_handles: *const u32,
        item_vqt: *const opc_da_bindings::tagOPCITEMVQT,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.write_vqt(
                count,
                item_server_handles.into_com_array_ref(count)?,
                item_vqt.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCAsyncIO2_Impl for Group_Impl<T> {
    fn Read(
        &self,
        count: u32,
        item_server_handles: *const u32,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.read2(
                item_server_handles.into_com_array_ref(count)?,
                transaction_id,
                cancel_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn Write(
        &self,
        count: u32,
        item_server_handles: *const u32,
        item_values: *const windows::Win32::System::Variant::VARIANT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.write2(
                count,
                item_server_handles.into_com_array_ref(count)?,
                item_values.into_com_array_ref(count)?,
                transaction_id,
                cancel_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn Refresh2(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        transaction_id: u32,
    ) -> windows::core::Result<u32> {
        self.refresh2(source, transaction_id)
    }

    fn Cancel2(&self, cancel_id: u32) -> windows::core::Result<()> {
        self.cancel2(cancel_id)
    }

    fn SetEnable(&self, enable: windows_core::BOOL) -> windows::core::Result<()> {
        self.set_enable(enable)
    }

    fn GetEnable(&self) -> windows::core::Result<windows_core::BOOL> {
        self.get_enable()
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCAsyncIO3_Impl for Group_Impl<T> {
    fn ReadMaxAge(
        &self,
        count: u32,
        item_server_handles: *const u32,
        max_age: *const u32,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.read_max_age2(
                item_server_handles.into_com_array_ref(count)?,
                max_age.into_com_array_ref(count)?,
                transaction_id,
                cancel_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn WriteVQT(
        &self,
        count: u32,
        item_server_handles: *const u32,
        item_vqt: *const opc_da_bindings::tagOPCITEMVQT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.write_vqt2(
                item_server_handles.into_com_array_ref(count)?,
                item_vqt.into_com_array_ref(count)?,
                transaction_id,
                cancel_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn RefreshMaxAge(&self, max_age: u32, transaction_id: u32) -> windows::core::Result<u32> {
        self.refresh_max_age(max_age, transaction_id)
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 required
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCItemDeadbandMgt_Impl for Group_Impl<T> {
    fn SetItemDeadband(
        &self,
        count: u32,
        item_server_handles: *const u32,
        percent_deadband: *const f32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_item_deadband(
                item_server_handles.into_com_array_ref(count)?,
                percent_deadband.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn GetItemDeadband(
        &self,
        count: u32,
        item_server_handles: *const u32,
        percent_deadband: *mut *mut f32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.get_item_deadband(
                item_server_handles.into_com_array_ref(count)?,
                percent_deadband.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            percent_deadband,
            errors
        }
    }

    fn ClearItemDeadband(
        &self,
        count: u32,
        item_server_handles: *const u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.clear_item_deadband(
                item_server_handles.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }
}

// 1.0 N/A
// 2.0 N/A
// 3.0 optional
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCItemSamplingMgt_Impl for Group_Impl<T> {
    fn SetItemSamplingRate(
        &self,
        count: u32,
        item_server_handles: *const u32,
        requested_sampling_rate: *const u32,
        revised_sampling_rate: *mut *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_item_sampling_rate(
                count,
                item_server_handles.into_com_array_ref(count)?,
                requested_sampling_rate.into_com_array_ref(count)?,
                revised_sampling_rate.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            revised_sampling_rate,
            errors
        }
    }

    fn GetItemSamplingRate(
        &self,
        count: u32,
        item_server_handles: *const u32,
        sampling_rate: *mut *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.get_item_sampling_rate(
                item_server_handles.into_com_array_ref(count)?,
                sampling_rate.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            sampling_rate,
            errors
        }
    }

    fn ClearItemSamplingRate(
        &self,
        count: u32,
        item_server_handles: *const u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.clear_item_sampling_rate(
                item_server_handles.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn SetItemBufferEnable(
        &self,
        count: u32,
        item_server_handles: *const u32,
        penable: *const windows_core::BOOL,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.set_item_buffer_enable(
                item_server_handles.into_com_array_ref(count)?,
                penable.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn GetItemBufferEnable(
        &self,
        count: u32,
        item_server_handles: *const u32,
        enable: *mut *mut windows_core::BOOL,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.get_item_buffer_enable(
                item_server_handles.into_com_array_ref(count)?,
                enable.into_com_array_ref(count)?,
                errors.into_com_array_ref(count)?,
            ),
            enable,
            errors
        }
    }
}

// 1.0 N/A
// 2.0 required
// 3.0 required
impl<T: GroupTrait + 'static> windows::Win32::System::Com::IConnectionPointContainer_Impl
    for Group_Impl<T>
{
    fn EnumConnectionPoints(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints> {
        self.enum_connection_points()
    }

    fn FindConnectionPoint(
        &self,
        reference_interface_id: *const windows::core::GUID,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint> {
        self.find_connection_point(reference_interface_id.into_ref()?)
    }
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> opc_da_bindings::IOPCAsyncIO_Impl for Group_Impl<T> {
    fn Read(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
        count: u32,
        item_server_handles: *const u32,
        transaction_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.read3(
                connection,
                source,
                item_server_handles.into_com_array_ref(count)?,
                transaction_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn Write(
        &self,
        connection: u32,
        count: u32,
        item_server_handles: *const u32,
        item_values: *const windows::Win32::System::Variant::VARIANT,
        transaction_id: *mut u32,
        errors: *mut *mut windows::core::HRESULT,
    ) -> windows::core::Result<()> {
        safe_call! {
            self.write3(
                connection,
                item_server_handles.into_com_array_ref(count)?,
                item_values.into_com_array_ref(count)?,
                transaction_id.into_ref()?,
                errors.into_com_array_ref(count)?,
            ),
            errors
        }
    }

    fn Refresh(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
    ) -> windows::core::Result<u32> {
        self.refresh(connection, source)
    }

    fn Cancel(&self, transaction_id: u32) -> windows::core::Result<()> {
        self.cancel(transaction_id)
    }
}

// 1.0 required
// 2.0 optional
// 3.0 N/A
impl<T: GroupTrait + 'static> windows::Win32::System::Com::IDataObject_Impl for Group_Impl<T> {
    fn GetData(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::Result<windows::Win32::System::Com::STGMEDIUM> {
        self.get_data(format_etc_in.into_ref()?)
    }

    fn GetDataHere(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
        storage_medium: *mut windows::Win32::System::Com::STGMEDIUM,
    ) -> windows::core::Result<()> {
        self.get_data_here(format_etc_in.into_ref()?, storage_medium.into_ref()?)
    }

    fn QueryGetData(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::HRESULT {
        let format_etc_in = match format_etc_in.into_ref() {
            Ok(format_etc_in) => format_etc_in,
            Err(err) => return err.code(),
        };

        self.query_get_data(format_etc_in)
    }

    fn GetCanonicalFormatEtc(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
        format_etc_inout: *mut windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::HRESULT {
        let format_etc_in = match format_etc_in.into_ref() {
            Ok(format_etc_in) => format_etc_in,
            Err(err) => return err.code(),
        };

        let format_etc_inout = match format_etc_inout.into_ref() {
            Ok(format_etc_inout) => format_etc_inout,
            Err(err) => return err.code(),
        };

        self.get_canonical_format_etc(format_etc_in, format_etc_inout)
    }

    fn SetData(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
        storage_medium: *const windows::Win32::System::Com::STGMEDIUM,
        release: windows_core::BOOL,
    ) -> windows::core::Result<()> {
        self.set_data(
            format_etc_in.into_ref()?,
            storage_medium.into_ref()?,
            release,
        )
    }

    fn EnumFormatEtc(
        &self,
        direction: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumFORMATETC> {
        self.enum_format_etc(direction)
    }

    fn DAdvise(
        &self,
        format_etc_in: *const windows::Win32::System::Com::FORMATETC,
        adv: u32,
        sink: windows::core::Ref<'_, windows::Win32::System::Com::IAdviseSink>,
    ) -> windows::core::Result<u32> {
        self.data_advise(format_etc_in.into_ref()?, adv, sink)
    }

    fn DUnadvise(&self, connection: u32) -> windows::core::Result<()> {
        self.data_unadvise(connection)
    }

    fn EnumDAdvise(&self) -> windows::core::Result<windows::Win32::System::Com::IEnumSTATDATA> {
        self.enum_data_advise()
    }
}
