pub fn get_class_id_from_program_id<
    T: TryInto<opc_classic_utils::CallerAllocatedWString, Error = windows::core::Error>,
>(
    program_id: T,
) -> windows_core::Result<windows_core::GUID> {
    let program_id: opc_classic_utils::CallerAllocatedWString = program_id.try_into()?;
    let id = unsafe { windows::Win32::System::Com::CLSIDFromProgID(program_id.as_pcwstr())? };
    Ok(id)
}
