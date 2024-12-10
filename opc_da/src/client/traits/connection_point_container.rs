use windows::core::GUID;
use windows::Win32::System::Com::IConnectionPoint;

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
    fn find_connection_point(&self, id: &GUID) -> windows::core::Result<IConnectionPoint> {
        unsafe { self.interface()?.FindConnectionPoint(id) }
    }

    /// Enumerates all available connection points.
    ///
    /// # Returns
    /// Enumerator for iterating through available connection points
    fn enum_connection_points(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints> {
        unsafe { self.interface()?.EnumConnectionPoints() }
    }
}
