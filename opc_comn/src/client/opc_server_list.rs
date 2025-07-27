use std::{mem::ManuallyDrop, str::FromStr as _};

use windows_core::Interface as _;

#[derive(Default)]
pub struct CreateOpcServerListOptions<'a> {
    pub class_id: Option<windows_core::GUID>,
    pub context: Option<windows::Win32::System::Com::CLSCTX>,
    pub outer: Option<&'a windows_core::IUnknown>,
    pub server_info: Option<&'a windows::Win32::System::Com::COSERVERINFO>,
}

pub struct OpcServerList {
    pub inner: opc_comn_bindings::IOPCServerList,
}

impl OpcServerList {
    pub fn create_opc_server_list(
        options: Option<CreateOpcServerListOptions>,
    ) -> windows_core::Result<OpcServerList> {
        let options = options.unwrap_or_default();

        let class_id = match options.class_id {
            Some(class_id) => class_id,
            None => {
                super::utils::get_class_id_from_program_id(windows_core::w!("OPC.ServerList.1"))?
            }
        };

        let context = options
            .context
            .unwrap_or(windows::Win32::System::Com::CLSCTX_ALL);

        let outer = options.outer;

        let server_list: opc_comn_bindings::IOPCServerList = unsafe {
            match options.server_info {
                Some(info) => {
                    let mut results = [windows::Win32::System::Com::MULTI_QI {
                        pIID: &opc_comn_bindings::IOPCServerList::IID,
                        ..Default::default()
                    }];

                    windows::Win32::System::Com::CoCreateInstanceEx(
                        &class_id,
                        outer,
                        context,
                        Some(info),
                        &mut results,
                    )?;

                    let result = results.into_iter().next().unwrap();

                    if result.hr.is_err() {
                        return Err(result.hr.into());
                    }

                    ManuallyDrop::into_inner(result.pItf)
                        .ok_or(windows::Win32::Foundation::E_POINTER)?
                        .cast::<opc_comn_bindings::IOPCServerList>()?
                }
                None => windows::Win32::System::Com::CoCreateInstance(&class_id, outer, context)?,
            }
        };

        Ok(OpcServerList { inner: server_list })
    }
}

pub struct ClassDetails {
    pub program_id: Option<String>,
    pub user_type: Option<String>,
}

impl OpcServerList {
    pub fn enum_classes_of_categories(
        &self,
        implemented_categories: &[windows_core::GUID],
        required_categories: &[windows_core::GUID],
    ) -> windows_core::Result<windows::Win32::System::Com::IEnumGUID> {
        unsafe {
            self.inner
                .EnumClassesOfCategories(implemented_categories, required_categories)
        }
    }

    pub fn get_class_details(
        &self,
        clsid: &windows_core::GUID,
    ) -> windows_core::Result<ClassDetails> {
        let mut program_id = opc_classic_utils::CalleeAllocatedWString::null();
        let mut user_type = opc_classic_utils::CalleeAllocatedWString::null();

        unsafe {
            self.inner.GetClassDetails(
                clsid,
                program_id.as_pwstr_mut_ptr(),
                user_type.as_pwstr_mut_ptr(),
            )?;

            Ok(ClassDetails {
                program_id: program_id.to_string()?,
                user_type: user_type.to_string()?,
            })
        }
    }

    pub fn class_id_from_program_id(
        &self,
        program_id: &str,
    ) -> windows_core::Result<windows_core::GUID> {
        unsafe {
            self.inner.CLSIDFromProgID(
                opc_classic_utils::CallerAllocatedWString::from_str(program_id)?.as_pcwstr(),
            )
        }
    }
}
