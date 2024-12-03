/// Iterator over COM GUIDs from IEnumGUID.  
///
/// # Safety  
/// This struct wraps a COM interface and must be used according to COM rules.  
pub struct GuidIter {
    iter: windows::Win32::System::Com::IEnumGUID,
    cache: [windows_core::GUID; 16],
    count: u32,
    finished: bool,
}

impl GuidIter {
    /// Creates a new iterator from a COM interface.  
    pub(super) fn new(iter: windows::Win32::System::Com::IEnumGUID) -> Self {
        Self {
            iter,
            cache: [windows_core::GUID::zeroed(); 16],
            count: 0,
            finished: false,
        }
    }
}

impl Iterator for GuidIter {
    type Item = windows_core::Result<windows_core::GUID>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }

        if self.count == 0 {
            if self.count > 0 {
                self.count -= 1;
                return Some(Ok(self.cache[self.count as usize]));
            }

            let ids = &mut self.cache;
            let count = &mut self.count;

            let code = unsafe { self.iter.Next(ids, Some(count)) };
            if code.is_err() {
                self.finished = true;
                return Some(Err(windows_core::Error::new(
                    code,
                    "Failed to get next GUID",
                )));
            }
        }

        if self.count == 0 {
            self.finished = true;
            None
        } else {
            self.count -= 1;
            Some(Ok(self.cache[self.count as usize]))
        }
    }
}
