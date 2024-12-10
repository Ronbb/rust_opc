use crate::client::RemotePointer;

/// Iterator over COM GUIDs from IEnumGUID.  
///
/// # Safety  
/// This struct wraps a COM interface and must be used according to COM rules.  
pub struct GuidIterator {
    iter: windows::Win32::System::Com::IEnumGUID,
    cache: [windows::core::GUID; 16],
    count: u32,
    finished: bool,
}

impl GuidIterator {
    /// Creates a new iterator from a COM interface.  
    pub fn new(iter: windows::Win32::System::Com::IEnumGUID) -> Self {
        Self {
            iter,
            cache: [windows::core::GUID::zeroed(); 16],
            count: 0,
            finished: false,
        }
    }
}

impl Iterator for GuidIterator {
    type Item = windows::core::Result<windows::core::GUID>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.count == 0 {
            let ids = &mut self.cache;
            let count = &mut self.count;

            let code = unsafe { self.iter.Next(ids, Some(count)) };
            if code.is_ok() {
                if *count == 0 {
                    self.finished = true;
                    return None;
                }
            } else {
                self.finished = true;
                return Some(Err(windows::core::Error::new(
                    code,
                    "Failed to get next GUID",
                )));
            }
        }

        self.count -= 1;
        assert!(self.count < self.cache.len() as u32);
        Some(Ok(self.cache[self.count as usize]))
    }
}

pub struct StringIterator {
    iter: windows::Win32::System::Com::IEnumString,
    cache: [windows::core::PWSTR; 16],
    count: u32,
    finished: bool,
}

impl StringIterator {
    pub fn new(iter: windows::Win32::System::Com::IEnumString) -> Self {
        Self {
            iter,
            cache: [windows::core::PWSTR::null(); 16],
            count: 0,
            finished: false,
        }
    }
}

impl Iterator for StringIterator {
    type Item = windows::core::Result<String>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.count == 0 {
            let ids = &mut self.cache;
            let count = &mut self.count;

            let code = unsafe { self.iter.Next(ids, Some(count)) };
            if code.is_ok() {
                if *count == 0 {
                    self.finished = true;
                    return None;
                }
            } else {
                self.finished = true;
                return Some(Err(windows::core::Error::new(
                    code,
                    "Failed to get next string",
                )));
            }
        }

        self.count -= 1;

        Some(RemotePointer::from(self.cache[self.count as usize]).try_into())
    }
}
