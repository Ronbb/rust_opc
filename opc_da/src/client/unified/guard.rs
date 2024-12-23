/// A RAII guard that manages COM initialization and uninitialization for a thread.
///
/// This type ensures that COM is properly initialized when the guard is created and
/// properly uninitialized when the guard is dropped. It wraps an inner value of type `T`
/// and provides access to it through the `Deref` trait.
///
/// # Thread Safety
/// The guard is intentionally not `Send` and not `Sync` to ensure COM operations
/// remain on the thread where they were initialized.
#[derive(Debug)]
pub struct Guard<T> {
    inner: T,
    /// Marker to ensure `Client` is not `Send` and not `Sync`.
    _marker: std::marker::PhantomData<*const ()>,
}

impl<T> Guard<T> {
    /// Creates a new guard that initializes COM and wraps the provided value.
    ///
    /// # Arguments
    /// * `value` - The value to wrap in the guard
    ///
    /// # Returns
    /// Returns a `Result` containing the guard if COM initialization succeeds.
    ///
    /// # Errors
    /// Returns an error if COM initialization fails.
    pub fn new(value: T) -> windows::core::Result<Self> {
        let guard = Self {
            inner: value,
            _marker: std::marker::PhantomData,
        };

        Self::try_initialize()?;

        Ok(guard)
    }
}

/// Provides direct access to the wrapped value through reference.
impl<T> std::ops::Deref for Guard<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.inner
    }
}

/// Ensures COM is uninitialized when the guard is dropped.
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
    pub(crate) fn try_initialize() -> windows::core::Result<()> {
        unsafe {
            windows::Win32::System::Com::CoInitializeEx(
                None,
                windows::Win32::System::Com::COINIT_MULTITHREADED,
            )
        }
        .ok()
    }

    /// Initializes COM for the current thread, panicking on failure.
    ///
    /// # Panics
    /// Panics if COM initialization fails.
    ///
    /// # Thread Safety
    /// COM initialization is performed with COINIT_MULTITHREADED flag.
    pub(crate) fn initialize() {
        Self::try_initialize().expect("Failed to initialize COM");
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
