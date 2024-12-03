use super::Client;

impl Client {
    /// Ensures COM is initialized for the current thread.
    ///
    /// # Returns
    /// Returns the HRESULT of the COM initialization.
    ///
    /// # Thread Safety
    /// COM initialization is performed with COINIT_MULTITHREADED flag.
    ///
    /// # Note
    /// Callers should check the returned HRESULT for initialization failures.
    pub(crate) fn initialize() -> windows_core::HRESULT {
        unsafe {
            windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_MULTITHREADED,
            )
        }
    }

    /// Uninitializes COM for the current thread.
    ///
    /// # Safety
    /// This method should be called when the thread is shutting down
    /// and no more COM calls will be made.
    pub(crate) fn uninitialize() {
        unsafe { windows::Win32::System::Com::CoUninitialize() };
    }
}
