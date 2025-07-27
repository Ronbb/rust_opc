use windows_core::*;
use opc_da_bindings::*;
use opc_classic_utils::*;

#[windows::core::implement(IOPCBrowse)]
pub struct OPCBrowse<T>(T)
where
    T: 'static + traits::OPCBrowse;

impl<T: traits::OPCBrowse> IOPCBrowse_Impl for OPCBrowse_Impl<T> {
    fn GetProperties(
        &self,
        dwitemcount: u32,
        pszitemids: *const PCWSTR,
        breturnpropertyvalues: BOOL,
        dwpropertycount: u32,
        pdwpropertyids: *const u32,
        ppitemproperties: *mut *mut tagOPCITEMPROPERTIES,
    ) -> Result<()> {
        let item_ids = unsafe {
            std::slice::from_raw_parts(pszitemids, dwitemcount as usize)
        };
        let property_ids = unsafe {
            std::slice::from_raw_parts(pdwpropertyids, dwpropertycount as usize)
        };
        
        let item_ids: Vec<String> = item_ids.iter().map(|pcwstr| {
            unsafe {
                CallerAllocatedWString::from_pcwstr(*pcwstr)
                    .to_string()
                    .unwrap_or_default()
            }
        }).collect();

        let properties = self.0.get_properties(item_ids, breturnpropertyvalues.as_bool(), property_ids.to_vec())?;
        
        // Allocate and write properties
        if !ppitemproperties.is_null() {
            let properties_array = CallerAllocatedArray::from_slice(&properties)?;
            unsafe { *ppitemproperties = properties_array.as_ptr() };
        }

        Ok(())
    }

    fn Browse(
        &self,
        szitemid: &PCWSTR,
        pszcontinuationpoint: *mut PWSTR,
        dwmaxelementsreturned: u32,
        dwbrowsefilter: tagOPCBROWSEFILTER,
        szelementnamefilter: &PCWSTR,
        szvendorfilter: &PCWSTR,
        breturnallproperties: BOOL,
        breturnpropertyvalues: BOOL,
        dwpropertycount: u32,
        pdwpropertyids: *const u32,
        pbmoreelements: *mut BOOL,
        pdwcount: *mut u32,
        ppbrowseelements: *mut *mut tagOPCBROWSEELEMENT,
    ) -> Result<()> {
        let item_id = unsafe {
            CallerAllocatedWString::from_pcwstr(*szitemid)
                .to_string()
                .unwrap_or_default()
        };
        let element_name_filter = unsafe {
            CallerAllocatedWString::from_pcwstr(*szelementnamefilter)
                .to_string()
                .unwrap_or_default()
        };
        let vendor_filter = unsafe {
            CallerAllocatedWString::from_pcwstr(*szvendorfilter)
                .to_string()
                .unwrap_or_default()
        };
        let property_ids = unsafe {
            std::slice::from_raw_parts(pdwpropertyids, dwpropertycount as usize)
        };

        let (more_elements, count, elements) = self.0.browse(
            item_id,
            dwmaxelementsreturned,
            dwbrowsefilter,
            element_name_filter,
            vendor_filter,
            breturnallproperties.as_bool(),
            breturnpropertyvalues.as_bool(),
            property_ids.to_vec(),
        )?;
        
        // Write output parameters
        if !pszcontinuationpoint.is_null() {
            unsafe { *pszcontinuationpoint = PWSTR::null() };
        }
        if !pbmoreelements.is_null() {
            unsafe { *pbmoreelements = BOOL::from(more_elements) };
        }
        if !pdwcount.is_null() {
            unsafe { *pdwcount = count };
        }
        if !ppbrowseelements.is_null() {
            let elements_array = CallerAllocatedArray::from_slice(&elements)?;
            unsafe { *ppbrowseelements = elements_array.as_ptr() };
        }

        Ok(())
    }
}

pub mod traits {
    use super::*;

    pub trait OPCBrowse {
        fn get_properties(&self, item_ids: Vec<String>, return_property_values: bool, property_ids: Vec<u32>) -> Result<Vec<tagOPCITEMPROPERTIES>>;
        
        fn browse(
            &self,
            item_id: String,
            max_elements_returned: u32,
            browse_filter: tagOPCBROWSEFILTER,
            element_name_filter: String,
            vendor_filter: String,
            return_all_properties: bool,
            return_property_values: bool,
            property_ids: Vec<u32>,
        ) -> Result<(bool, u32, Vec<tagOPCBROWSEELEMENT>)>;
    }
} 