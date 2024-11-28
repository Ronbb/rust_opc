pub trait ServerTrait {
    // IOPCCommon
    fn set_locale_id(&self, locale_id: u32) -> windows_core::Result<()>;

    // IOPCCommon
    fn get_locale_id(&self) -> windows_core::Result<u32>;

    // IOPCCommon
    fn query_available_locale_ids(&self) -> windows_core::Result<Vec<u32>>;

    // IOPCCommon
    fn get_error_string(&self, error: i32) -> windows_core::Result<String>;

    // IOPCCommon
    fn set_client_name(&self, name: String) -> windows_core::Result<()>;


    
}
