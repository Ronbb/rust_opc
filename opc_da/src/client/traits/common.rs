use crate::client::memory::{LocalPointer, RemotePointer};
use opc_da_bindings::IOPCCommon;

pub trait CommonTrait {
    fn interface(&self) -> &IOPCCommon;

    fn set_locale_id(&self, lcid: u32) -> windows::core::Result<()> {
        unsafe { self.interface().SetLocaleID(lcid) }
    }

    fn get_locale_id(&self) -> windows::core::Result<u32> {
        unsafe { self.interface().GetLocaleID() }
    }

    fn query_available_locale_ids(&self) -> windows::core::Result<Vec<u32>> {
        let mut count = 0;
        let mut lcids = RemotePointer::new();

        unsafe {
            self.interface()
                .QueryAvailableLocaleIDs(&mut count, lcids.as_mut_ptr())?;

            let slice = std::slice::from_raw_parts(lcids.as_option().unwrap(), count as usize);
            Ok(slice.to_vec())
        }
    }

    fn get_error_string(&self, error: windows::core::HRESULT) -> windows::core::Result<String> {
        let output = unsafe { self.interface().GetErrorString(error)? };

        RemotePointer::from(output).try_into()
    }

    fn set_client_name(&self, name: &str) -> windows::core::Result<()> {
        let name = LocalPointer::from(name);
        unsafe { self.interface().SetClientName(name.as_pcwstr()) }
    }
}
