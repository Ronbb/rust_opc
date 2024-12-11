pub trait GroupTrait {
    fn add_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
        results: &mut [opc_da_bindings::tagOPCITEMRESULT],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn validate_items(
        &self,
        items: &[opc_da_bindings::tagOPCITEMDEF],
        blob_update: windows::Win32::Foundation::BOOL,
        validation_results: &mut [opc_da_bindings::tagOPCITEMRESULT],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn remove_items(
        &self,
        item_server_handles: &[u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn set_active_state(
        &self,
        item_server_handles: &[u32],
        active: windows::Win32::Foundation::BOOL,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn set_client_handles(
        &self,
        item_server_handles: &[u32],
        handle_client: &[u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn set_data_types(
        &self,
        item_server_handles: &[u32],
        requested_data_types: &[u16],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn create_enumerator(
        &self,
        reference_interface_id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown>;

    #[allow(clippy::too_many_arguments)]
    fn get_state(
        &self,
        update_rate: &mut u32,
        active: &mut windows::Win32::Foundation::BOOL,
        name: &mut windows::core::PWSTR,
        time_bias: &mut i32,
        percent_deadband: &mut f32,
        locale_id: &mut u32,
        group_client_handle: &mut u32,
        item_server_handles_group: &mut u32,
    ) -> windows::core::Result<()>;

    #[allow(clippy::too_many_arguments)]
    fn set_state(
        &self,
        requested_update_rate: &u32,
        revised_update_rate: &mut u32,
        active: &windows::Win32::Foundation::BOOL,
        time_bias: &i32,
        percent_deadband: &f32,
        locale_id: &u32,
        group_client_handle: &u32,
    ) -> windows::core::Result<()>;

    fn set_name(&self, name: &windows::core::PCWSTR) -> windows::core::Result<()>;

    fn clone_group(
        &self,
        name: &windows::core::PCWSTR,
        reference_interface_id: &windows::core::GUID,
    ) -> windows::core::Result<windows::core::IUnknown>;

    fn set_keep_alive(&self, keep_alive_time: u32) -> windows::core::Result<u32>;

    fn get_keep_alive(&self) -> windows::core::Result<u32>;

    fn get_public_group_state(&self) -> windows::core::Result<windows::Win32::Foundation::BOOL>;

    fn move_to_public(&self) -> windows::core::Result<()>;

    fn read(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        item_server_handles: &[u32],
        item_values: &mut [opc_da_bindings::tagOPCITEMSTATE],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn write(
        &self,
        item_server_handles: &[u32],
        item_values: &[windows::core::VARIANT],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn read_max_age(
        &self,
        item_server_handles: &[u32],
        max_age: &[u32],
        values: &mut [windows::core::VARIANT],
        qualities: &mut [u16],
        timestamps: &mut [windows::Win32::Foundation::FILETIME],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn write_vqt(
        &self,
        count: u32,
        item_server_handles: &[u32],
        item_vqt: &[opc_da_bindings::tagOPCITEMVQT],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn read2(
        &self,
        item_server_handles: &[u32],
        transaction_id: u32,
        cancel_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn write2(
        &self,
        count: u32,
        item_server_handles: &[u32],
        item_values: &[windows::core::VARIANT],
        transaction_id: u32,
        cancel_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn refresh2(
        &self,
        source: opc_da_bindings::tagOPCDATASOURCE,
        transaction_id: u32,
    ) -> windows::core::Result<u32>;

    fn cancel2(&self, cancel_id: u32) -> windows::core::Result<()>;

    fn set_enable(&self, enable: windows::Win32::Foundation::BOOL) -> windows::core::Result<()>;

    fn get_enable(&self) -> windows::core::Result<windows::Win32::Foundation::BOOL>;

    fn read_max_age2(
        &self,
        item_server_handles: &[u32],
        max_age: &[u32],
        transaction_id: u32,
        cancel_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn write_vqt2(
        &self,
        item_server_handles: &[u32],
        item_vqt: &[opc_da_bindings::tagOPCITEMVQT],
        transaction_id: u32,
        cancel_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn refresh_max_age(&self, max_age: u32, transaction_id: u32) -> windows::core::Result<u32>;

    fn set_item_deadband(
        &self,
        item_server_handles: &[u32],
        percent_deadband: &[f32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn get_item_deadband(
        &self,
        item_server_handles: &[u32],
        percent_deadband: &mut [f32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn clear_item_deadband(
        &self,
        item_server_handles: &[u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn set_item_sampling_rate(
        &self,
        count: u32,
        item_server_handles: &[u32],
        requested_sampling_rate: &[u32],
        revised_sampling_rate: &mut [u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn get_item_sampling_rate(
        &self,
        item_server_handles: &[u32],
        sampling_rate: &mut [u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn clear_item_sampling_rate(
        &self,
        item_server_handles: &[u32],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn set_item_buffer_enable(
        &self,
        item_server_handles: &[u32],
        penable: &windows::Win32::Foundation::BOOL,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn get_item_buffer_enable(
        &self,
        item_server_handles: &[u32],
        enable: &mut [windows::Win32::Foundation::BOOL],
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn enum_connection_points(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints>;

    fn find_connection_point(
        &self,
        reference_interface_id: &windows::core::GUID,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint>;

    fn read3(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
        item_server_handles: &[u32],
        transaction_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn write3(
        &self,
        connection: u32,
        item_server_handles: &[u32],
        item_values: &[windows::core::VARIANT],
        transaction_id: &mut u32,
        errors: &mut [windows::core::HRESULT],
    ) -> windows::core::Result<()>;

    fn refresh(
        &self,
        connection: u32,
        source: opc_da_bindings::tagOPCDATASOURCE,
    ) -> windows::core::Result<u32>;

    fn cancel(&self, transaction_id: u32) -> windows::core::Result<()>;

    fn get_data(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::Result<windows::Win32::System::Com::STGMEDIUM>;

    fn get_data_here(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
        storage_medium: &mut windows::Win32::System::Com::STGMEDIUM,
    ) -> windows::core::Result<()>;

    fn query_get_data(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::HRESULT;

    fn get_canonical_format_etc(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
        format_etc_out: &mut windows::Win32::System::Com::FORMATETC,
    ) -> windows::core::HRESULT;

    fn set_data(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
        medium: &windows::Win32::System::Com::STGMEDIUM,
        release: windows::Win32::Foundation::BOOL,
    ) -> windows::core::Result<()>;

    fn enum_format_etc(
        &self,
        direction: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumFORMATETC>;

    /// Establishes an advisory connection.  
    ///  
    /// # Arguments  
    /// * `sink` - The sink interface. If None, any existing connection will be removed.  
    ///  
    /// # Returns  
    /// * The connection token if sink is Some and connection is established  
    /// * Ok(0) if sink is None (indicating no connection)  
    /// * An error if the operation fails  
    fn data_advise(
        &self,
        format_etc_in: &windows::Win32::System::Com::FORMATETC,
        adv: u32,
        sink: Option<&windows::Win32::System::Com::IAdviseSink>,
    ) -> windows::core::Result<u32>;

    fn data_unadvise(&self, connection: u32) -> windows::core::Result<()>;

    fn enum_data_advise(&self) -> windows::core::Result<windows::Win32::System::Com::IEnumSTATDATA>;
}
