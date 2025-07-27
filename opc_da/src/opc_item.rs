use opc_classic_utils::*;
use opc_da_bindings::*;
use windows_core::*;

#[windows::core::implement(IOPCItemMgt)]
pub struct OPCItem<T>(T)
where
    T: 'static + traits::OPCItem;

impl<T: traits::OPCItem> IOPCItemMgt_Impl for OPCItem_Impl<T> {
    fn AddItems(
        &self,
        dwcount: u32,
        pitemarray: *const tagOPCITEMDEF,
        ppaddresults: *mut *mut tagOPCITEMRESULT,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let items = unsafe { std::slice::from_raw_parts(pitemarray, dwcount as usize) };

        let item_defs: Vec<traits::ItemDefinition> = items
            .iter()
            .map(|item| traits::ItemDefinition {
                access_path: unsafe {
                    CallerAllocatedWString::from_pwstr(item.szAccessPath)
                        .to_string_lossy()
                        .unwrap_or_default()
                },
                item_id: unsafe {
                    CallerAllocatedWString::from_pwstr(item.szItemID)
                        .to_string_lossy()
                        .unwrap_or_default()
                },
                active: item.bActive.as_bool(),
                client_handle: item.hClient,
                blob_size: item.dwBlobSize,
                blob: if !item.pBlob.is_null() {
                    unsafe {
                        std::slice::from_raw_parts(item.pBlob, item.dwBlobSize as usize).to_vec()
                    }
                } else {
                    Vec::new()
                },
                requested_data_type: item.vtRequestedDataType,
            })
            .collect();

        let (results, errors) = self.0.add_items(item_defs)?;

        // Allocate and write results
        copy_to_caller_array!(ppaddresults, &results);

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors)?;

        Ok(())
    }

    fn ValidateItems(
        &self,
        dwcount: u32,
        pitemarray: *const tagOPCITEMDEF,
        bblobupdate: BOOL,
        ppvalidationresults: *mut *mut tagOPCITEMRESULT,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let items = unsafe { std::slice::from_raw_parts(pitemarray, dwcount as usize) };

        let item_defs: Vec<traits::ItemDefinition> = items
            .iter()
            .map(|item| traits::ItemDefinition {
                access_path: unsafe {
                    if item.szAccessPath.is_null() {
                        String::new()
                    } else {
                        CallerAllocatedWString::from_pcwstr(PCWSTR(item.szAccessPath.0))
                            .to_string_lossy()
                            .unwrap_or_default()
                    }
                },
                item_id: unsafe {
                    if item.szItemID.is_null() {
                        String::new()
                    } else {
                        CallerAllocatedWString::from_pcwstr(PCWSTR(item.szItemID.0))
                            .to_string_lossy()
                            .unwrap_or_default()
                    }
                },
                active: item.bActive.as_bool(),
                client_handle: item.hClient,
                blob_size: item.dwBlobSize,
                blob: if !item.pBlob.is_null() {
                    unsafe {
                        std::slice::from_raw_parts(item.pBlob, item.dwBlobSize as usize).to_vec()
                    }
                } else {
                    Vec::new()
                },
                requested_data_type: item.vtRequestedDataType,
            })
            .collect();

        let (results, errors) = self.0.validate_items(item_defs, bblobupdate.as_bool())?;

        // Allocate and write results
        copy_to_caller_array!(ppvalidationresults, &results);

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors);

        Ok(())
    }

    fn RemoveItems(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };

        let errors = self.0.remove_items(server_handles.to_vec())?;

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors);

        Ok(())
    }

    fn SetActiveState(
        &self,
        dwcount: u32,
        phserver: *const u32,
        bactive: BOOL,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };

        let errors = self
            .0
            .set_active_state(server_handles.to_vec(), bactive.as_bool())?;

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors);

        Ok(())
    }

    fn SetClientHandles(
        &self,
        dwcount: u32,
        phserver: *const u32,
        phclient: *const u32,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };
        let client_handles = unsafe { std::slice::from_raw_parts(phclient, dwcount as usize) };

        let errors = self
            .0
            .set_client_handles(server_handles.to_vec(), client_handles.to_vec())?;

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors);

        Ok(())
    }

    fn SetDatatypes(
        &self,
        dwcount: u32,
        phserver: *const u32,
        pvrequesteddatatypes: *const u16,
        pperrors: *mut *mut HRESULT,
    ) -> Result<()> {
        let server_handles = unsafe { std::slice::from_raw_parts(phserver, dwcount as usize) };
        let requested_datatypes =
            unsafe { std::slice::from_raw_parts(pvrequesteddatatypes, dwcount as usize) };

        let errors = self
            .0
            .set_datatypes(server_handles.to_vec(), requested_datatypes.to_vec())?;

        // Allocate and write errors
        copy_to_caller_array!(pperrors, &errors);

        Ok(())
    }

    fn CreateEnumerator(&self, riid: *const GUID) -> Result<IUnknown> {
        self.0.create_enumerator(riid)
    }
}

pub mod traits {
    use super::*;

    #[derive(Debug, Clone)]
    pub struct ItemDefinition {
        pub access_path: String,
        pub item_id: String,
        pub active: bool,
        pub client_handle: u32,
        pub blob_size: u32,
        pub blob: Vec<u8>,
        pub requested_data_type: u16,
    }

    pub trait OPCItem {
        fn add_items(
            &self,
            items: Vec<ItemDefinition>,
        ) -> Result<(Vec<tagOPCITEMRESULT>, Vec<HRESULT>)>;

        fn validate_items(
            &self,
            items: Vec<ItemDefinition>,
            blob_update: bool,
        ) -> Result<(Vec<tagOPCITEMRESULT>, Vec<HRESULT>)>;

        fn remove_items(&self, server_handles: Vec<u32>) -> Result<Vec<HRESULT>>;

        fn set_active_state(&self, server_handles: Vec<u32>, active: bool) -> Result<Vec<HRESULT>>;

        fn set_client_handles(
            &self,
            server_handles: Vec<u32>,
            client_handles: Vec<u32>,
        ) -> Result<Vec<HRESULT>>;

        fn set_datatypes(
            &self,
            server_handles: Vec<u32>,
            requested_datatypes: Vec<u16>,
        ) -> Result<Vec<HRESULT>>;

        fn create_enumerator(&self, riid: *const GUID) -> Result<IUnknown>;
    }
}
