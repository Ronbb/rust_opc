pub struct GuidIter(pub(super) windows::Win32::System::Com::IEnumGUID);

impl Iterator for GuidIter {
    type Item = windows_core::GUID;

    fn next(&mut self) -> Option<Self::Item> {
        let mut ids = [windows_core::GUID::zeroed(); 1];
        let mut count = ids.len() as u32;

        unsafe { self.0.Next(&mut ids, Some(&mut count)).ok() }.ok()?;

        if count == 0 {
            None
        } else {
            Some(ids[0])
        }
    }
}
