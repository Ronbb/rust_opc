use windows::Win32::System::Com::{FORMATETC, STGMEDIUM};

pub trait DataObjectTrait {
    fn interface(&self) -> windows_core::Result<&windows::Win32::System::Com::IDataObject>;

    fn get_data(&self, format: &FORMATETC) -> windows::core::Result<STGMEDIUM> {
        unsafe { self.interface()?.GetData(format) }
    }

    fn get_data_here(&self, format: &FORMATETC) -> windows::core::Result<STGMEDIUM> {
        let mut output = STGMEDIUM::default();
        unsafe { self.interface()?.GetDataHere(format, &mut output)? };
        Ok(output)
    }

    fn query_get_data(&self, format: &FORMATETC) -> windows::core::Result<()> {
        unsafe { self.interface()?.QueryGetData(format) }.ok()
    }

    fn get_canonical_format(&self, format_in: &FORMATETC) -> windows::core::Result<FORMATETC> {
        let mut output = FORMATETC::default();
        unsafe {
            self.interface()?
                .GetCanonicalFormatEtc(format_in, &mut output)
        }
        .ok()?;

        Ok(output)
    }

    fn set_data(
        &self,
        format: &FORMATETC,
        medium: &STGMEDIUM,
        release: bool,
    ) -> windows::core::Result<()> {
        unsafe {
            self.interface()?.SetData(
                format,
                medium,
                windows::Win32::Foundation::BOOL::from(release),
            )
        }
    }

    fn enumerate_formats(
        &self,
        direction: u32,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumFORMATETC> {
        unsafe { self.interface()?.EnumFormatEtc(direction) }
    }

    fn dadvise(
        &self,
        format: &FORMATETC,
        advf: u32,
        sink: &windows::Win32::System::Com::IAdviseSink,
    ) -> windows::core::Result<u32> {
        unsafe { self.interface()?.DAdvise(format, advf, sink) }
    }

    fn dunadvise(&self, connection: u32) -> windows::core::Result<()> {
        unsafe { self.interface()?.DUnadvise(connection) }
    }

    fn enum_dadvise(&self) -> windows::core::Result<windows::Win32::System::Com::IEnumSTATDATA> {
        unsafe { self.interface()?.EnumDAdvise() }
    }
}
