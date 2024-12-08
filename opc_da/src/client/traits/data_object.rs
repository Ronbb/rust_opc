use windows::Win32::System::Com::{FORMATETC, STGMEDIUM};

pub trait DataObjectTrait {
    fn data_object(&self) -> &windows::Win32::System::Com::IDataObject;

    fn get_data(&self, format: &FORMATETC) -> windows::core::Result<STGMEDIUM> {
        unsafe { self.data_object().GetData(format) }
    }

    fn get_data_here(
        &self,
        format: &FORMATETC,
        medium: *mut STGMEDIUM,
    ) -> windows::core::Result<()> {
        unsafe { self.data_object().GetDataHere(format, medium) }
    }

    fn query_get_data(&self, format: &FORMATETC) -> windows::core::Result<()> {
        unsafe { self.data_object().QueryGetData(format) }.ok()
    }

    fn get_canonical_format(
        &self,
        format_in: &FORMATETC,
        format_out: *mut FORMATETC,
    ) -> windows::core::Result<()> {
        unsafe {
            self.data_object()
                .GetCanonicalFormatEtc(format_in, format_out)
        }
        .ok()
    }

    fn set_data(
        &self,
        format: &FORMATETC,
        medium: &STGMEDIUM,
        release: bool,
    ) -> windows::core::Result<()> {
        unsafe {
            self.data_object().SetData(
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
        unsafe { self.data_object().EnumFormatEtc(direction) }
    }

    fn dadvise(
        &self,
        format: &FORMATETC,
        advf: u32,
        sink: &windows::Win32::System::Com::IAdviseSink,
    ) -> windows::core::Result<u32> {
        unsafe { self.data_object().DAdvise(format, advf, sink) }
    }

    fn dunadvise(&self, connection: u32) -> windows::core::Result<()> {
        unsafe { self.data_object().DUnadvise(connection) }
    }

    fn enum_dadvise(&self) -> windows::core::Result<windows::Win32::System::Com::IEnumSTATDATA> {
        unsafe { self.data_object().EnumDAdvise() }
    }
}
