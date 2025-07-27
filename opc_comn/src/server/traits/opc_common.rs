pub trait OpcCommon {
    fn set_locale_id(&self, dwlcid: u32) -> windows_core::Result<()>;
    fn get_locale_id(&self) -> windows_core::Result<u32>;
    fn query_available_locale_ids(&self) -> windows_core::Result<Vec<u32>>;
    fn get_error_string(&self, dwerror: windows_core::HRESULT) -> windows_core::Result<String>;
    fn set_client_name(&self, szname: String) -> windows_core::Result<()>;
}
