use opc_da_bindings::IOPCCommon;

use crate::utils::{LocalPointer, RemoteArray, RemotePointer};

/// Common OPC server functionality trait.
///
/// Provides methods for locale management and error string retrieval.
/// This trait is implemented by all OPC DA servers to support basic
/// configuration and error handling capabilities.
pub trait CommonTrait {
    fn interface(&self) -> windows::core::Result<&IOPCCommon>;

    /// Sets the locale ID for server string localization.
    ///
    /// # Arguments
    /// * `locale_id` - Windows LCID (Locale ID) value for the desired language
    ///
    /// # Returns
    /// Result indicating if the locale was successfully set
    ///
    /// # Examples
    /// ```no_run
    /// use windows::Win32::System::SystemServices::LOCALE_USER_DEFAULT;
    /// server.set_locale_id(LOCALE_USER_DEFAULT);
    /// ```
    fn set_locale_id(&self, locale_id: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.SetLocaleID(locale_id) }
    }

    /// Gets the current locale ID used by the server.
    ///
    /// # Returns
    /// Windows LCID value representing the current locale
    fn get_locale_id(&self) -> windows::core::Result<u32> {
        unsafe { self.interface()?.GetLocaleID() }
    }

    /// Gets a list of locale IDs supported by the server.
    ///
    /// # Returns
    /// Array of Windows LCID values for supported locales
    fn query_available_locale_ids(&self) -> windows::core::Result<RemoteArray<u32>> {
        let mut locale_ids = RemoteArray::empty();

        unsafe {
            self.interface()?
                .QueryAvailableLocaleIDs(locale_ids.as_mut_len_ptr(), locale_ids.as_mut_ptr())?;
        }

        Ok(locale_ids)
    }

    /// Gets a localized error description string.
    ///
    /// # Arguments
    /// * `error` - HRESULT error code to get description for
    ///
    /// # Returns
    /// Localized error message string in current locale
    fn get_error_string(&self, error: windows::core::HRESULT) -> windows::core::Result<String> {
        let output = unsafe { self.interface()?.GetErrorString(error)? };

        RemotePointer::from(output).try_into()
    }

    /// Sets a client name for server identification.
    ///
    /// # Arguments
    /// * `name` - Client application name or description
    ///
    /// # Returns
    /// Result indicating if the client name was successfully set
    fn set_client_name(&self, name: &str) -> windows::core::Result<()> {
        let name = LocalPointer::from(name);
        unsafe { self.interface()?.SetClientName(name.as_pcwstr()) }
    }
}
