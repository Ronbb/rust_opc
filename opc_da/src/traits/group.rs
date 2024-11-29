use windows::Win32::System::Com::{
    IAdviseSink, IConnectionPoint, IEnumConnectionPoints, IEnumFORMATETC, IEnumSTATDATA,
};

use super::def::*;

pub trait GroupTrait {
    fn add_items(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn validate_items(
        &self,
        count: u32,
        item_array: *const bindings::tagOPCITEMDEF,
        blob_update: bool,
        validation_results: *mut *mut bindings::tagOPCITEMRESULT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn remove_items(
        &self,
        count: u32,
        handle_server: Option<u32>,
    ) -> windows_core::Result<Vec<windows_core::HRESULT>>;

    fn set_active_state(
        &self,
        count: u32,
        handle_server: Option<u32>,
        active: bool,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn set_client_handles(
        &self,
        count: u32,
        handle_server: Option<u32>,
        handle_client: Option<u32>,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn set_datatypes(
        &self,
        count: u32,
        handle_server: Option<u32>,
        requested_data_types: Option<u16>,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn create_enumerator(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown>;

    fn get_state(
        &self,
        update_rate: *mut u32,
        active: *mut bool,
        name: *mut windows_core::PWSTR,
        timebias: *mut i32,
        percent_deadband: *mut f32,
        locale_id: *mut u32,
        handle_client_group: *mut u32,
        handle_server_group: *mut u32,
    ) -> windows_core::Result<()>;

    fn set_state(
        &self,
        requested_update_rate: Option<u32>,
        revised_update_rate: *mut u32,
        active: Option<bool>,
        timebias: *const i32,
        percent_deadband: *const f32,
        locale_id: Option<u32>,
        handle_client_group: Option<u32>,
    ) -> windows_core::Result<()>;

    fn set_name(&self, name: String) -> windows_core::Result<()>;

    fn clone_group(
        &self,
        name: String,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<windows_core::IUnknown>;

    fn set_keep_alive(&self, keep_alive_time: u32) -> windows_core::Result<u32>;

    fn get_keep_alive(&self) -> windows_core::Result<u32>;

    fn get_state2(&self) -> windows_core::Result<bool>;

    fn move_to_public(&self) -> windows_core::Result<()>;

    fn read(
        &self,
        source: bindings::tagOPCDATASOURCE,
        count: u32,
        handle_server: Option<u32>,
        item_values: *mut *mut bindings::tagOPCITEMSTATE,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn write(
        &self,
        count: u32,
        handle_server: Option<u32>,
        item_values: *const windows_core::VARIANT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn read_max_age(
        &self,
        count: u32,
        handle_server: Option<u32>,
        max_age: Option<u32>,
        values: *mut *mut windows_core::VARIANT,
        qualities: *mut *mut u16,
        timestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn write_vqt(
        &self,
        count: u32,
        handle_server: Option<u32>,
        item_vqt: *const bindings::tagOPCITEMVQT,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn read2(
        &self,
        count: u32,
        handle_server: Option<u32>,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn write2(
        &self,
        count: u32,
        handle_server: Option<u32>,
        item_values: *const windows_core::VARIANT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn refresh2(
        &self,
        source: bindings::tagOPCDATASOURCE,
        transaction_id: u32,
    ) -> windows_core::Result<u32>;

    fn cancel2(&self, cancel_id: u32) -> windows_core::Result<()>;

    fn set_enable(&self, enable: bool) -> windows_core::Result<()>;

    fn get_enable(&self) -> windows_core::Result<bool>;

    fn read_max_age2(
        &self,
        count: u32,
        handle_server: Option<u32>,
        max_age: Option<u32>,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn write_vqt2(
        &self,
        count: u32,
        handle_server: Option<u32>,
        item_vqt: *const bindings::tagOPCITEMVQT,
        transaction_id: u32,
        cancel_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn refresh_max_age(&self, max_age: u32, transaction_id: u32) -> windows_core::Result<u32>;

    fn set_item_deadband(
        &self,
        count: u32,
        handle_server: Option<u32>,
        percent_deadband: *const f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn get_item_deadband(
        &self,
        count: u32,
        handle_server: Option<u32>,
        percent_deadband: *mut *mut f32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn clear_item_deadband(
        &self,
        count: u32,
        handle_server: Option<u32>,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn set_item_sampling_rate(
        &self,
        count: u32,
        handle_server: Option<u32>,
        requested_sampling_rate: Option<u32>,
        revised_sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn get_item_sampling_rate(
        &self,
        count: u32,
        handle_server: Option<u32>,
        sampling_rate: *mut *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn clear_item_sampling_rate(
        &self,
        count: u32,
        handle_server: Option<u32>,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn set_item_buffer_enable(
        &self,
        count: u32,
        handle_server: Option<u32>,
        penable: Option<bool>,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn get_item_buffer_enable(
        &self,
        count: u32,
        handle_server: Option<u32>,
        enable: *mut *mut bool,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn enum_connection_points(&self) -> windows_core::Result<IEnumConnectionPoints>;

    fn find_connection_point(
        &self,
        reference_interface_id: *const windows_core::GUID,
    ) -> windows_core::Result<IConnectionPoint>;

    fn read3(
        &self,
        connection: u32,
        source: DataSource,
        count: u32,
        handle_server: Option<u32>,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn write3(
        &self,
        connection: u32,
        count: u32,
        handle_server: *const u32,
        item_values: *const windows_core::VARIANT,
        transaction_id: *mut u32,
        errors: *mut *mut windows_core::HRESULT,
    ) -> windows_core::Result<()>;

    fn refresh(&self, connection: u32, source: DataSource) -> windows_core::Result<u32>;

    fn cancel(&self, transaction_id: u32) -> windows_core::Result<()>;

    fn get_data(&self, format_etc_in: Option<FormatEtc>) -> windows_core::Result<StorageMedium>;

    fn get_data_here(
        &self,
        format_etc_in: Option<FormatEtc>,
    ) -> windows_core::Result<StorageMedium>;

    fn query_get_data(&self, format_etc_in: Option<FormatEtc>) -> windows_core::HRESULT;

    fn get_canonical_format_etc(
        &self,
        format_etc_in: Option<FormatEtc>,
    ) -> windows_core::Result<FormatEtc>;

    fn set_data(
        &self,
        format_etc_in: Option<FormatEtc>,
        medium: Option<StorageMedium>,
        release: bool,
    ) -> windows_core::Result<()>;

    fn enum_format_etc(&self, direction: u32) -> windows_core::Result<IEnumFORMATETC>;

    fn data_advise(
        &self,
        format_etc_in: Option<FormatEtc>,
        adv: u32,
        sink: Option<&IAdviseSink>,
    ) -> windows_core::Result<u32>;

    fn data_unadvise(&self, connection: u32) -> windows_core::Result<()>;

    fn enum_data_advise(&self) -> windows_core::Result<IEnumSTATDATA>;
}
