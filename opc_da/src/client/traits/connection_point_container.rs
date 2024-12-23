use windows::core::{Interface as _, GUID};

/// COM connection point container functionality.
///
/// Provides methods to establish connections between event sources
/// and event sinks in the OPC COM architecture. Used primarily for
/// handling asynchronous callbacks.
pub trait ConnectionPointContainerTrait {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer>;

    /// Finds a connection point for a specific interface.
    ///
    /// # Arguments
    /// * `id` - GUID of the connection point interface to find
    ///
    /// # Returns
    /// Connection point interface for the specified GUID
    ///
    /// # Example
    /// ```no_run
    /// use windows::core::GUID;
    /// let cp = container.find_connection_point(&GUID_CALLBACK_INTERFACE);
    /// ```
    ///
    /// # Safety  
    /// Caller must ensure:  
    /// - COM is properly initialized  
    /// - The underlying COM object is valid  
    ///
    /// # Errors  
    /// Returns an error if:  
    /// - The COM operation fails  
    /// - The connection point is not found
    fn find_connection_point(
        &self,
        id: &GUID,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint> {
        unsafe { self.interface()?.FindConnectionPoint(id) }
    }

    fn data_callback_connection_point(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IConnectionPoint> {
        self.find_connection_point(&opc_da_bindings::IOPCDataCallback::IID)
    }

    /// Enumerates all available connection points.
    ///
    /// # Returns
    /// Enumerator for iterating through available connection points
    ///
    /// # Safety  
    /// Caller must ensure:  
    /// - COM is properly initialized  
    /// - The underlying COM object is valid  
    ///
    /// # Errors  
    /// Returns an error if:  
    /// - The COM operation fails  
    /// - No connection points are available  
    fn enum_connection_points(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints> {
        unsafe { self.interface()?.EnumConnectionPoints() }
    }
}
