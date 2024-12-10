#[derive(Debug)]
pub struct Guard<T> {
    inner: T,
    /// Marker to ensure `Client` is not `Send` and not `Sync`.
    _marker: std::marker::PhantomData<*const ()>,
}

impl<T> Guard<T> {
    pub fn new(value: T) -> windows::core::Result<Self> {
        let guard = Self {
            inner: value,
            _marker: std::marker::PhantomData,
        };

        Self::try_initialize().ok()?;

        Ok(guard)
    }
}

impl<T> std::ops::Deref for Guard<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.inner
    }
}

impl<T> Drop for Guard<T> {
    fn drop(&mut self) {
        Self::uninitialize();
    }
}

impl<T> Guard<T> {
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
    pub(crate) fn try_initialize() -> windows::core::HRESULT {
        unsafe {
            windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_MULTITHREADED,
            )
        }
    }

    pub(crate) fn initialize() {
        Self::try_initialize()
            .ok()
            .expect("Failed to initialize COM");
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
