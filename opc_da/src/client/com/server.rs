use windows_core::Interface as _;

pub struct Server {
    pub(crate) raw: opc_da_bindings::IOPCServer,
}

impl std::ops::Deref for Server {
    type Target = opc_da_bindings::IOPCServer;

    fn deref(&self) -> &Self::Target {
        &self.raw
    }
}

impl Server {
    #[allow(clippy::too_many_arguments)]
    pub fn add_group(
        &self,
        name: &str,
        active: bool,
        client_handle: u32,
        update_rate: u32,
        locale_id: u32,
        time_bias: i32,
        percent_deadband: f32,
    ) -> windows_core::Result<opc_da_bindings::IOPCItemMgt> {
        let mut group = None;
        let mut group_server_handle = 0u32;
        let mut group_name = name.encode_utf16().chain(Some(0)).collect::<Vec<_>>();
        let group_name = windows_core::PWSTR::from_raw(group_name.as_mut_ptr());
        let mut revised_percent_deadband = 0;

        unsafe {
            self.raw.AddGroup(
                group_name,
                windows::Win32::Foundation::BOOL::from(active),
                update_rate,
                client_handle,
                &time_bias,
                &percent_deadband,
                locale_id,
                &mut group_server_handle,
                &mut revised_percent_deadband,
                &opc_da_bindings::IOPCItemMgt::IID,
                &mut group,
            )?;

            match group {
                None => Err(windows_core::Error::new(
                    windows::Win32::Foundation::E_POINTER,
                    "Failed to add group, returned null",
                )),
                Some(group) => Ok(group.cast()?),
            }
        }
    }
}
