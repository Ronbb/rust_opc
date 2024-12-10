use crate::client::{
    memory::{LocalPointer, RemotePointer},
    RemoteArray,
};
use opc_da_bindings::IOPCCommon;

pub trait CommonTrait {
    fn interface(&self) -> windows::core::Result<&IOPCCommon>;

    fn set_locale_id(&self, locale_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.SetLocaleID(locale_id) }
    }

    fn get_locale_id(&self) -> windows::core::Result<u32> {
        unsafe { self.interface()?.GetLocaleID() }
    }

    fn query_available_locale_ids(&self) -> windows::core::Result<RemoteArray<u32>> {
        let mut locale_ids = RemoteArray::empty();

        unsafe {
            self.interface()?
                .QueryAvailableLocaleIDs(locale_ids.as_mut_len_ptr(), locale_ids.as_mut_ptr())?;
        }

        Ok(locale_ids)
    }

    fn get_error_string(&self, error: windows::core::HRESULT) -> windows::core::Result<String> {
        let output = unsafe { self.interface()?.GetErrorString(error)? };

        RemotePointer::from(output).try_into()
    }

    fn set_client_name(&self, name: &str) -> windows::core::Result<()> {
        let name = LocalPointer::from(name);
        unsafe { self.interface()?.SetClientName(name.as_pcwstr()) }
    }
}
