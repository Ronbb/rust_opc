use windows::Win32::System::Com::{FORMATETC, STGMEDIUM};

/// Data transfer functionality using COM's structured storage.
///
/// Provides methods to transfer data between client and server using
/// structured storage formats and advisory connections.
pub trait DataObjectTrait {
    fn interface(&self) -> windows::core::Result<&windows::Win32::System::Com::IDataObject>;

    /// Gets data from the object in the specified format.
    ///
    /// # Arguments
    /// * `format` - Format specification including clipboard format and storage medium
    ///
    /// # Returns
    /// Storage medium containing the requested data
    fn get_data(&self, format: &FORMATETC) -> windows::core::Result<STGMEDIUM> {
        unsafe { self.interface()?.GetData(format) }
    }

    /// Gets data in place using the specified format.
    ///
    /// # Arguments
    /// * `format` - Format specification including clipboard format and storage medium
    ///
    /// # Returns
    /// Storage medium updated with the requested data
    fn get_data_here(&self, format: &FORMATETC) -> windows::core::Result<STGMEDIUM> {
        let mut output = STGMEDIUM::default();
        unsafe { self.interface()?.GetDataHere(format, &mut output)? };
        Ok(output)
    }

    /// Tests if data is available in the specified format.
    ///
    /// # Arguments
    /// * `format` - Format specification to test for availability
    ///
    /// # Returns
    /// Ok(()) if the format is supported, error otherwise
    fn query_get_data(&self, format: &FORMATETC) -> windows::core::Result<()> {
        unsafe { self.interface()?.QueryGetData(format) }.ok()
    }

    /// Gets the canonical format equivalent to the specified format.
    ///
    /// # Arguments
    /// * `format_in` - Format specification to convert
    ///
    /// # Returns
    /// Canonical format specification
    fn get_canonical_format(&self, format_in: &FORMATETC) -> windows::core::Result<FORMATETC> {
        let mut output = FORMATETC::default();
        unsafe {
            self.interface()?
                .GetCanonicalFormatEtc(format_in, &mut output)
        }
        .ok()?;

        Ok(output)
    }

    /// Sets data in the specified format.
    ///
    /// # Arguments
    /// * `format` - Format specification for the data
    /// * `medium` - Storage medium containing the data
    /// * `release` - If true, the object takes ownership of medium
    ///
    /// # Returns
    /// Ok(()) if data was set successfully
    fn set_data(
        &self,
        format: &FORMATETC,
        medium: &STGMEDIUM,
        release: bool,
    ) -> windows::core::Result<()> {
        unsafe { self.interface()?.SetData(format, medium, release) }
    }

    /// Enumerates available data formats.
    ///
    /// # Arguments
    /// * `direction` - Direction of data flow (DATADIR_GET = 1, DATADIR_SET = 2)
    ///
    /// # Returns
    /// Enumerator for available format specifications
    fn enumerate_formats(
        &self,
        direction: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumFORMATETC> {
        unsafe { self.interface()?.EnumFormatEtc(direction) }
    }

    /// Establishes an advisory connection for data change notifications.
    ///
    /// # Arguments
    /// * `format` - Format specification to monitor
    /// * `advf` - Advisory flags controlling notification behavior
    /// * `sink` - Sink interface to receive notifications
    ///
    /// # Returns
    /// Connection token for the advisory connection
    fn dadvise(
        &self,
        format: &FORMATETC,
        advf: u32,
        sink: &windows::Win32::System::Com::IAdviseSink,
    ) -> windows::core::Result<u32> {
        unsafe { self.interface()?.DAdvise(format, advf, sink) }
    }

    /// Terminates an advisory connection.
    ///
    /// # Arguments
    /// * `connection` - Connection token from dadvise
    ///
    /// # Returns
    /// Ok(()) if connection was terminated successfully
    fn dunadvise(&self, connection: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.DUnadvise(connection) }
    }

    /// Enumerates active advisory connections.
    ///
    /// # Returns
    /// Enumerator for active advisory connections
    fn enum_dadvise(&self) -> windows::core::Result<windows::Win32::System::Com::IEnumSTATDATA> {
        unsafe { self.interface()?.EnumDAdvise() }
    }
}
