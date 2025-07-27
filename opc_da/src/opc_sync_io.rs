use opc_classic_utils::*;
use opc_da_bindings::*;
use windows_core::*;

#[windows::core::implement(IOPCSyncIO, IOPCSyncIO2)]
pub struct OPCSyncIO<T>(T)
where
    T: 'static + traits::OPCSyncIO + traits::OPCSyncIO2;

impl<T: traits::OPCSyncIO + traits::OPCSyncIO2> IOPCSyncIO_Impl for OPCSyncIO_Impl<T> {
    fn Read(
        &self,
        dwsource: tagOPCDATASOURCE,
        dwcount: u32,
        phserver: *const u32,
        ppitemvalues: *mut *mut tagOPCITEMSTATE,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };

        let (item_states, errors) = self.0.read(dwsource, server_handles.to_vec())?;

        // Allocate and write item states
        if !ppitemvalues.is_null() {
            let size = std::mem::size_of::<tagOPCITEMSTATE>() * item_states.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(item_states.as_ptr(), ptr.cast(), item_states.len());
                *ppitemvalues = ptr.cast();
            }
        }

        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(errors.as_ptr(), ptr.cast(), errors.len());
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn Write(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pitemvalues: *const windows::Win32::System::Variant::VARIANT,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };
        let values = unsafe { std::slice::from_raw_parts(pitemvalues, dwcount as usize) };

        let errors = self.0.write(server_handles.to_vec(), values.to_vec())?;

        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(errors.as_ptr(), ptr.cast(), errors.len());
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }
}

impl<T: traits::OPCSyncIO + traits::OPCSyncIO2> IOPCSyncIO2_Impl for OPCSyncIO_Impl<T> {
    fn ReadMaxAge(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pdwmaxage: *const u32,
        ppvvalues: *mut *mut windows::Win32::System::Variant::VARIANT,
        ppwqualities: *mut *mut u16,
        ppfttimestamps: *mut *mut windows::Win32::Foundation::FILETIME,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };
        let max_ages = unsafe { std::slice::from_raw_parts(pdwmaxage, dwcount as usize) };

        let (values, qualities, timestamps, errors) = self
            .0
            .read_max_age(server_handles.to_vec(), max_ages.to_vec())?;

        // Allocate and write values
        if !ppvvalues.is_null() {
            let size =
                std::mem::size_of::<windows::Win32::System::Variant::VARIANT>() * values.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(values.as_ptr(), ptr.cast(), values.len());
                *ppvvalues = ptr.cast();
            }
        }

        // Allocate and write qualities
        if !ppwqualities.is_null() {
            let size = std::mem::size_of::<u16>() * qualities.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(qualities.as_ptr(), ptr.cast(), qualities.len());
                *ppwqualities = ptr.cast();
            }
        }

        // Allocate and write timestamps
        if !ppfttimestamps.is_null() {
            let size =
                std::mem::size_of::<windows::Win32::Foundation::FILETIME>() * timestamps.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(timestamps.as_ptr(), ptr.cast(), timestamps.len());
                *ppfttimestamps = ptr.cast();
            }
        }

        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(errors.as_ptr(), ptr.cast(), errors.len());
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn WriteVQT(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pitemvqt: *const tagOPCITEMVQT,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };
        let item_vqt = unsafe { std::slice::from_raw_parts(pitemvqt, dwcount as usize) };

        let errors = self
            .0
            .write_vqt(server_handles.to_vec(), item_vqt.to_vec())?;

        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }

            unsafe {
                std::ptr::copy_nonoverlapping(errors.as_ptr(), ptr.cast(), errors.len());
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }
}

pub mod traits {
    use super::*;

    pub trait OPCSyncIO {
        fn read(
            &self,
            source: tagOPCDATASOURCE,
            server_handles: Vec<u32>,
        ) -> Result<(Vec<tagOPCITEMSTATE>, Vec<HRESULT>)>;

        fn write(
            &self,
            server_handles: Vec<u32>,
            values: Vec<windows::Win32::System::Variant::VARIANT>,
        ) -> Result<Vec<HRESULT>>;
    }

    pub trait OPCSyncIO2: OPCSyncIO {
        fn read_max_age(
            &self,
            server_handles: Vec<u32>,
            max_ages: Vec<u32>,
        ) -> Result<(
            Vec<windows::Win32::System::Variant::VARIANT>,
            Vec<u16>,
            Vec<windows::Win32::Foundation::FILETIME>,
            Vec<HRESULT>,
        )>;

        fn write_vqt(
            &self,
            server_handles: Vec<u32>,
            item_vqt: Vec<tagOPCITEMVQT>,
        ) -> Result<Vec<HRESULT>>;
    }
}
