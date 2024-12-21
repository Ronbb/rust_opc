use crate::{
    def::ItemAttributes,
    utils::{RemoteArray, RemotePointer, TryToLocal as _},
};

const MAX_CACHE_SIZE: usize = 16;

/// Iterator over COM GUIDs from IEnumGUID.  
///
/// # Safety  
/// This struct wraps a COM interface and must be used according to COM rules.  
pub struct GuidIterator {
    inner: windows::Win32::System::Com::IEnumGUID,
    cache: Box<[windows::core::GUID; MAX_CACHE_SIZE]>,
    index: u32,
    count: u32,
    done: bool,
}

impl GuidIterator {
    /// Creates a new iterator from a COM interface.  
    pub fn new(inner: windows::Win32::System::Com::IEnumGUID) -> Self {
        Self {
            inner,
            cache: Box::from([windows::core::GUID::zeroed(); MAX_CACHE_SIZE]),
            index: 0,
            count: 0,
            done: false,
        }
    }
}

impl Iterator for GuidIterator {
    type Item = windows::core::Result<windows::core::GUID>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        if self.index == self.cache.len() as u32 {
            let code = unsafe {
                self.inner
                    .Next(self.cache.as_mut_slice(), Some(&mut self.count))
            };

            if code.is_ok() {
                if self.count == 0 {
                    self.done = true;
                    return None;
                }

                self.index = 0;
            } else {
                self.done = true;
                return Some(Err(windows::core::Error::new(
                    code,
                    "Failed to get next GUID",
                )));
            }
        }

        let current = self.cache[self.index as usize];
        self.index += 1;
        Some(Ok(current))
    }
}

pub struct StringIterator {
    inner: windows::Win32::System::Com::IEnumString,
    cache: Box<[windows::core::PWSTR; MAX_CACHE_SIZE]>,
    index: u32,
    count: u32,
    done: bool,
}

impl StringIterator {
    pub fn new(inner: windows::Win32::System::Com::IEnumString) -> Self {
        Self {
            inner,
            cache: Box::new([windows::core::PWSTR::null(); MAX_CACHE_SIZE]),
            index: 0,
            count: 0,
            done: false,
        }
    }
}

impl Iterator for StringIterator {
    type Item = windows::core::Result<String>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        if self.index == self.cache.len() as u32 {
            let code = unsafe {
                self.inner
                    .Next(self.cache.as_mut_slice(), Some(&mut self.count))
            };

            if code.is_ok() {
                if self.count == 0 {
                    self.done = true;
                    return None;
                }

                self.index = 0;
            } else {
                self.done = true;
                return Some(Err(windows::core::Error::new(
                    code,
                    "Failed to get next string",
                )));
            }
        }

        let current = RemotePointer::from(self.cache[self.index as usize]);
        self.index += 1;
        Some(current.try_into())
    }
}

pub struct GroupIterator<Group: TryFrom<windows::core::IUnknown, Error = windows::core::Error>> {
    inner: windows::Win32::System::Com::IEnumUnknown,
    cache: Box<[Option<windows::core::IUnknown>; MAX_CACHE_SIZE]>,
    index: u32,
    count: u32,
    done: bool,
    _mark: std::marker::PhantomData<Group>,
}

impl<Group: TryFrom<windows::core::IUnknown, Error = windows::core::Error>> GroupIterator<Group> {
    pub fn new(inner: windows::Win32::System::Com::IEnumUnknown) -> Self {
        Self {
            inner,
            cache: Box::from([const { None }; MAX_CACHE_SIZE]),
            index: 0,
            count: 0,
            done: false,
            _mark: std::marker::PhantomData,
        }
    }
}

impl<Group: TryFrom<windows::core::IUnknown, Error = windows::core::Error>> Iterator
    for GroupIterator<Group>
{
    type Item = windows::core::Result<Group>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        if self.index == self.cache.len() as u32 {
            let code = unsafe {
                self.inner
                    .Next(self.cache.as_mut_slice(), Some(&mut self.count))
            };

            if code.is_ok() {
                if self.count == 0 {
                    self.done = true;
                    return None;
                }

                self.index = 0;
            } else {
                self.done = true;
                return Some(Err(windows::core::Error::new(
                    code,
                    "Failed to get next group",
                )));
            }
        }

        let current = self.cache[self.index as usize].take();
        self.index += 1;
        Some(match current {
            Some(group) => group.try_into(),
            None => Err(windows::core::Error::new(
                windows::Win32::Foundation::E_POINTER,
                "Failed to get group, returned null",
            )),
        })
    }
}

// for opc_da_bindings::IEnumOPCItemAttributes
pub struct ItemAttributeIterator {
    inner: opc_da_bindings::IEnumOPCItemAttributes,
    cache: RemoteArray<opc_da_bindings::tagOPCITEMATTRIBUTES>,
    index: u32,
    done: bool,
}

impl ItemAttributeIterator {
    pub fn new(inner: opc_da_bindings::IEnumOPCItemAttributes) -> Self {
        Self {
            inner,
            cache: RemoteArray::empty(),
            index: 0,
            done: false,
        }
    }
}

impl Iterator for ItemAttributeIterator {
    type Item = windows::core::Result<ItemAttributes>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        if self.index == self.cache.len() {
            let mut attrs = RemoteArray::new(MAX_CACHE_SIZE as u32);

            let result = unsafe {
                self.inner.Next(
                    MAX_CACHE_SIZE as u32,
                    attrs.as_mut_ptr(),
                    attrs.as_mut_len_ptr(),
                )
            };

            match result {
                Ok(_) => {
                    if attrs.is_empty() {
                        self.done = true;
                        return None;
                    }

                    self.cache = attrs;
                    self.index = 0;
                }
                Err(err) => {
                    self.done = true;
                    return Some(Err(err));
                }
            }
        }

        let current = self.cache.as_slice()[self.index as usize].try_to_local();
        self.index += 1;
        Some(current)
    }
}
