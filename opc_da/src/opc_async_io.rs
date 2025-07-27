use windows_core::*;
use opc_da_bindings::*;
use opc_classic_utils::*;

#[windows::core::implement(IOPCAsyncIO)]
pub struct OPCAsyncIO<T>(T)
where
    T: 'static + traits::OPCAsyncIO;

impl<T: traits::OPCAsyncIO> IOPCAsyncIO_Impl for OPCAsyncIO_Impl<T> {
    fn Read(
        &self,
        dwconnection: u32,
        dwsource: tagOPCDATASOURCE,
        dwcount: u32,
        phserver: *const u32,
        ptransactionid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        
        let (transaction_id, errors) = self.0.read(dwconnection, dwsource, server_handles.to_vec())?;
        
        write_caller_allocated_ptr!(ptransactionid, transaction_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn Write(
        &self,
        dwconnection: u32,
        dwcount: u32,
        phserver: *const u32,
        pitemvalues: *const windows::Win32::System::Variant::VARIANT,
        ptransactionid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        let values = unsafe {
            std::slice::from_raw_parts(pitemvalues, dwcount as usize)
        };

        let (transaction_id, errors) = self.0.write(dwconnection, server_handles.to_vec(), values.to_vec())?;
        
        write_caller_allocated_ptr!(ptransactionid, transaction_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn Refresh(
        &self,
        dwconnection: u32,
        dwsource: tagOPCDATASOURCE,
    ) -> Result<u32> {
        self.0.refresh(dwconnection, dwsource)
    }

    fn Cancel(&self, dwtransactionid: u32) -> Result<()> {
        self.0.cancel(dwtransactionid)
    }
}



#[windows::core::implement(IOPCAsyncIO2, IOPCAsyncIO3)]
pub struct OPCAsyncIO3<T>(T)
where
    T: 'static + traits::OPCAsyncIO3;

impl<T: traits::OPCAsyncIO3> IOPCAsyncIO2_Impl for OPCAsyncIO3_Impl<T> {
    fn Read(
        &self,
        dwcount: u32,
        phserver: *const u32,
        dwtransactionid: u32,
        pdwcancelid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        
        let (cancel_id, errors) = traits::OPCAsyncIO2::read(&self.0, dwcount, server_handles.to_vec(), dwtransactionid)?;
        
        write_caller_allocated_ptr!(pdwcancelid, cancel_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
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
        dwtransactionid: u32,
        pdwcancelid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        let values = unsafe {
            std::slice::from_raw_parts(pitemvalues, dwcount as usize)
        };

        let (cancel_id, errors) = traits::OPCAsyncIO2::write(&self.0, dwcount, server_handles.to_vec(), values.to_vec(), dwtransactionid)?;
        
        write_caller_allocated_ptr!(pdwcancelid, cancel_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn Refresh2(
        &self,
        dwsource: tagOPCDATASOURCE,
        dwtransactionid: u32,
    ) -> Result<u32> {
        self.0.refresh2(dwsource, dwtransactionid)
    }

    fn Cancel2(&self, dwcancelid: u32) -> Result<()> {
        self.0.cancel2(dwcancelid)
    }

    fn SetEnable(&self, benable: BOOL) -> Result<()> {
        self.0.set_enable(benable.as_bool())
    }

    fn GetEnable(&self) -> Result<BOOL> {
        let enabled = self.0.get_enable()?;
        Ok(BOOL::from(enabled))
    }
}

impl<T: traits::OPCAsyncIO3> IOPCAsyncIO3_Impl for OPCAsyncIO3_Impl<T> {
    fn ReadMaxAge(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pdwmaxage: *const u32,
        dwtransactionid: u32,
        pdwcancelid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        let max_ages = unsafe {
            std::slice::from_raw_parts(pdwmaxage, dwcount as usize)
        };
        
        let (cancel_id, errors) = self.0.read_max_age(dwcount, server_handles.to_vec(), max_ages.to_vec(), dwtransactionid)?;
        
        write_caller_allocated_ptr!(pdwcancelid, cancel_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
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
        dwtransactionid: u32,
        pdwcancelid: *mut u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe {
            std::slice::from_raw_parts(phserver, dwcount as usize)
        };
        let item_vqt = unsafe {
            std::slice::from_raw_parts(pitemvqt, dwcount as usize)
        };

        let (cancel_id, errors) = self.0.write_vqt(dwcount, server_handles.to_vec(), item_vqt.to_vec(), dwtransactionid)?;
        
        write_caller_allocated_ptr!(pdwcancelid, cancel_id)?;
        
        // Allocate and write errors
        if !pperrors.is_null() {
            let size = std::mem::size_of::<HRESULT>() * errors.len();
            let ptr = unsafe { windows::Win32::System::Com::CoTaskMemAlloc(size) };
            if ptr.is_null() {
                return Err(windows::core::Error::from_win32());
            }
            
            unsafe {
                std::ptr::copy_nonoverlapping(
                    errors.as_ptr(),
                    ptr.cast(),
                    errors.len()
                );
                *pperrors = ptr.cast();
            }
        }

        Ok(())
    }

    fn RefreshMaxAge(
        &self,
        dwmaxage: u32,
        dwtransactionid: u32,
    ) -> Result<u32> {
        self.0.refresh_max_age(dwmaxage, dwtransactionid)
    }
}

pub mod traits {
    use super::*;

    pub trait OPCAsyncIO {
        fn read(&self, connection: u32, source: tagOPCDATASOURCE, server_handles: Vec<u32>) -> Result<(u32, Vec<HRESULT>)>;
        
        fn write(&self, connection: u32, server_handles: Vec<u32>, values: Vec<windows::Win32::System::Variant::VARIANT>) -> Result<(u32, Vec<HRESULT>)>;
        
        fn refresh(&self, connection: u32, source: tagOPCDATASOURCE) -> Result<u32>;
        
        fn cancel(&self, transaction_id: u32) -> Result<()>;
    }

    pub trait OPCAsyncIO2: OPCAsyncIO {
        fn read(&self, count: u32, server_handles: Vec<u32>, transaction_id: u32) -> Result<(u32, Vec<HRESULT>)>;
        
        fn write(&self, count: u32, server_handles: Vec<u32>, values: Vec<windows::Win32::System::Variant::VARIANT>, transaction_id: u32) -> Result<(u32, Vec<HRESULT>)>;
        
        fn refresh2(&self, source: tagOPCDATASOURCE, transaction_id: u32) -> Result<u32>;
        
        fn cancel2(&self, cancel_id: u32) -> Result<()>;
        
        fn set_enable(&self, enable: bool) -> Result<()>;
        
        fn get_enable(&self) -> Result<bool>;
    }

    pub trait OPCAsyncIO3: OPCAsyncIO2 {
        fn read_max_age(&self, count: u32, server_handles: Vec<u32>, max_ages: Vec<u32>, transaction_id: u32) -> Result<(u32, Vec<HRESULT>)>;
        
        fn write_vqt(&self, count: u32, server_handles: Vec<u32>, item_vqt: Vec<tagOPCITEMVQT>, transaction_id: u32) -> Result<(u32, Vec<HRESULT>)>;
        
        fn refresh_max_age(&self, max_age: u32, transaction_id: u32) -> Result<u32>;
    }
} 