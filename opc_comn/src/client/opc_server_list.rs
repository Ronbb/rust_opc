use std::mem::ManuallyDrop;

use opc_comn_bindings::IOPCServerList;
use windows_core::Interface as _;

#[derive(Default)]
pub struct CreateOpcServerListOptions<'a> {
    pub class_id: Option<windows_core::GUID>,
    pub context: Option<windows::Win32::System::Com::CLSCTX>,
    pub outer: Option<&'a windows_core::IUnknown>,
    pub server_info: Option<&'a windows::Win32::System::Com::COSERVERINFO>,
}

pub struct OpcServerList {
    pub inner: IOPCServerList,
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

        let server_list: IOPCServerList = unsafe {
            match options.server_info {
                Some(info) => {
                    let mut results = [windows::Win32::System::Com::MULTI_QI {
                        pIID: &IOPCServerList::IID,
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
                        .cast::<IOPCServerList>()?
                }
                None => windows::Win32::System::Com::CoCreateInstance(&class_id, outer, context)?,
            }
        };

        Ok(OpcServerList { inner: server_list })
    }
}
