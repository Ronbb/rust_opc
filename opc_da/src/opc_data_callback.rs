use opc_da_bindings::*;
use windows_core::*;

#[windows::core::implement(IOPCDataCallback)]
pub struct OPCDataCallback<T>(T)
where
    T: 'static + traits::OPCDataCallback;

impl<T: traits::OPCDataCallback> IOPCDataCallback_Impl for OPCDataCallback_Impl<T> {
    fn OnDataChange(
        &self,
        dwtransid: u32,
        hgroup: u32,
        hrquality: HRESULT,
        hrerror: HRESULT,
        dwcount: u32,
        phclientitems: *const u32,
        pvvalues: *const windows::Win32::System::Variant::VARIANT,
        pwqualities: *const u16,
        pfttimestamps: *const windows::Win32::Foundation::FILETIME,
        perrors: *const HRESULT,
    ) -> Result<()> {
        let client_items = unsafe { std::slice::from_raw_parts(phclientitems, dwcount as usize) };
        let values = unsafe { std::slice::from_raw_parts(pvvalues, dwcount as usize) };
        let qualities = unsafe { std::slice::from_raw_parts(pwqualities, dwcount as usize) };
        let timestamps = unsafe { std::slice::from_raw_parts(pfttimestamps, dwcount as usize) };
        let errors = unsafe { std::slice::from_raw_parts(perrors, dwcount as usize) };

        self.0.on_data_change(
            dwtransid,
            hgroup,
            hrquality,
            hrerror,
            client_items.to_vec(),
            values.to_vec(),
            qualities.to_vec(),
            timestamps.to_vec(),
            errors.to_vec(),
        )
    }

    fn OnReadComplete(
        &self,
        dwtransid: u32,
        hgroup: u32,
        hrquality: HRESULT,
        hrerror: HRESULT,
        dwcount: u32,
        phclientitems: *const u32,
        pvvalues: *const windows::Win32::System::Variant::VARIANT,
        pwqualities: *const u16,
        pfttimestamps: *const windows::Win32::Foundation::FILETIME,
        perrors: *const HRESULT,
    ) -> Result<()> {
        let client_items = unsafe { std::slice::from_raw_parts(phclientitems, dwcount as usize) };
        let values = unsafe { std::slice::from_raw_parts(pvvalues, dwcount as usize) };
        let qualities = unsafe { std::slice::from_raw_parts(pwqualities, dwcount as usize) };
        let timestamps = unsafe { std::slice::from_raw_parts(pfttimestamps, dwcount as usize) };
        let errors = unsafe { std::slice::from_raw_parts(perrors, dwcount as usize) };

        self.0.on_read_complete(
            dwtransid,
            hgroup,
            hrquality,
            hrerror,
            client_items.to_vec(),
            values.to_vec(),
            qualities.to_vec(),
            timestamps.to_vec(),
            errors.to_vec(),
        )
    }

    fn OnWriteComplete(
        &self,
        dwtransid: u32,
        hgroup: u32,
        hrmastererr: HRESULT,
        dwcount: u32,
        pclienthandles: *const u32,
        perrors: *const HRESULT,
    ) -> Result<()> {
        let client_items = unsafe { std::slice::from_raw_parts(pclienthandles, dwcount as usize) };
        let errors = unsafe { std::slice::from_raw_parts(perrors, dwcount as usize) };

        self.0.on_write_complete(
            dwtransid,
            hgroup,
            hrmastererr,
            client_items.to_vec(),
            errors.to_vec(),
        )
    }

    fn OnCancelComplete(&self, dwtransid: u32, hgroup: u32) -> Result<()> {
        self.0.on_cancel_complete(dwtransid, hgroup)
    }
}

pub mod traits {
    use super::*;

    pub trait OPCDataCallback {
        fn on_data_change(
            &self,
            transaction_id: u32,
            group: u32,
            quality: HRESULT,
            error: HRESULT,
            client_items: Vec<u32>,
            values: Vec<windows::Win32::System::Variant::VARIANT>,
            qualities: Vec<u16>,
            timestamps: Vec<windows::Win32::Foundation::FILETIME>,
            errors: Vec<HRESULT>,
        ) -> Result<()>;

        fn on_read_complete(
            &self,
            transaction_id: u32,
            group: u32,
            quality: HRESULT,
            error: HRESULT,
            client_items: Vec<u32>,
            values: Vec<windows::Win32::System::Variant::VARIANT>,
            qualities: Vec<u16>,
            timestamps: Vec<windows::Win32::Foundation::FILETIME>,
            errors: Vec<HRESULT>,
        ) -> Result<()>;

        fn on_write_complete(
            &self,
            transaction_id: u32,
            group: u32,
            master_error: HRESULT,
            client_items: Vec<u32>,
            errors: Vec<HRESULT>,
        ) -> Result<()>;

        fn on_cancel_complete(&self, transaction_id: u32, group: u32) -> Result<()>;
    }
}
