//! Safe, owning client wrappers for OPC Common interfaces.

use opc_classic_utils::{CoTaskMemOut, NoCleanup, OwnedPwstr, WideCString};
use windows::Win32::Foundation::E_UNEXPECTED;
use windows_core::{Error, GUID, Result};

use crate::{IOPCCommon, IOPCEnumGUID, IOPCServerList2, IOPCShutdown};

#[derive(Clone)]
pub struct CommonClient {
    inner: IOPCCommon,
}

impl CommonClient {
    pub fn new(inner: IOPCCommon) -> Self {
        Self { inner }
    }

    pub fn set_locale(&self, locale: u32) -> Result<()> {
        unsafe { self.inner.SetLocaleID(locale) }
    }

    pub fn locale(&self) -> Result<u32> {
        unsafe { self.inner.GetLocaleID() }
    }

    pub fn available_locales(&self) -> Result<Vec<u32>> {
        let mut count = 0u32;
        let mut output = CoTaskMemOut::<u32>::new();
        let call = unsafe {
            self.inner
                .QueryAvailableLocaleIDs(&mut count, output.as_mut_ptr())
        };
        // SAFETY: The IDL declares a task-allocated array with `count` elements.
        let output = unsafe { output.into_array(count as usize, NoCleanup) }?;
        call?;
        Ok(output.as_slice().to_vec())
    }

    pub fn error_string(&self, error: windows_core::HRESULT) -> Result<String> {
        let value = unsafe { self.inner.GetErrorString(error) }?;
        // SAFETY: OPC Common transfers a task-allocated string to the caller.
        Ok(unsafe { OwnedPwstr::from_raw(value.0) }.to_string_lossy())
    }

    pub fn set_client_name(&self, name: &str) -> Result<()> {
        let name = WideCString::try_from(name)
            .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))?;
        unsafe { self.inner.SetClientName(name.as_pcwstr()) }
    }

    pub fn as_raw(&self) -> &IOPCCommon {
        &self.inner
    }
}

#[derive(Clone)]
pub struct GuidEnumerator {
    inner: IOPCEnumGUID,
}

impl GuidEnumerator {
    pub fn new(inner: IOPCEnumGUID) -> Self {
        Self { inner }
    }

    pub fn next_batch(&self, maximum: usize) -> Result<Vec<GUID>> {
        let capacity = u32::try_from(maximum)
            .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))?;
        let mut values = vec![GUID::zeroed(); capacity as usize];
        let mut fetched = 0u32;
        unsafe { self.inner.Next(&mut values, &mut fetched) }?;
        if fetched > capacity {
            return Err(Error::from_hresult(E_UNEXPECTED));
        }
        values.truncate(fetched as usize);
        Ok(values)
    }

    pub fn skip(&self, count: u32) -> Result<()> {
        unsafe { self.inner.Skip(count) }
    }

    pub fn reset(&self) -> Result<()> {
        unsafe { self.inner.Reset() }
    }

    pub fn try_clone(&self) -> Result<Self> {
        unsafe { self.inner.Clone() }.map(Self::new)
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ClassDetails {
    pub prog_id: String,
    pub user_type: String,
    pub version_independent_prog_id: String,
}

#[derive(Clone)]
pub struct ServerListClient {
    inner: IOPCServerList2,
}

impl ServerListClient {
    pub fn new(inner: IOPCServerList2) -> Self {
        Self { inner }
    }

    pub fn enum_classes(&self, implemented: &[GUID], required: &[GUID]) -> Result<GuidEnumerator> {
        unsafe { self.inner.EnumClassesOfCategories(implemented, required) }
            .map(GuidEnumerator::new)
    }

    pub fn class_details(&self, class_id: &GUID) -> Result<ClassDetails> {
        let mut prog_id = windows_core::PWSTR::null();
        let mut user_type = windows_core::PWSTR::null();
        let mut version_independent = windows_core::PWSTR::null();
        let call = unsafe {
            self.inner.GetClassDetails(
                class_id,
                &mut prog_id,
                &mut user_type,
                &mut version_independent,
            )
        };
        // Adopt every non-null output before propagating the HRESULT.
        let prog_id = unsafe { OwnedPwstr::from_raw(prog_id.0) };
        let user_type = unsafe { OwnedPwstr::from_raw(user_type.0) };
        let version_independent = unsafe { OwnedPwstr::from_raw(version_independent.0) };
        call?;
        Ok(ClassDetails {
            prog_id: prog_id.to_string_lossy(),
            user_type: user_type.to_string_lossy(),
            version_independent_prog_id: version_independent.to_string_lossy(),
        })
    }

    pub fn class_id_from_prog_id(&self, prog_id: &str) -> Result<GUID> {
        let prog_id = WideCString::try_from(prog_id)
            .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))?;
        unsafe { self.inner.CLSIDFromProgID(prog_id.as_pcwstr()) }
    }
}

#[derive(Clone)]
pub struct ShutdownClient {
    inner: IOPCShutdown,
}

impl ShutdownClient {
    pub fn new(inner: IOPCShutdown) -> Self {
        Self { inner }
    }

    pub fn request(&self, reason: &str) -> Result<()> {
        let reason = WideCString::try_from(reason)
            .map_err(|_| Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))?;
        unsafe { self.inner.ShutdownRequest(reason.as_pcwstr()) }
    }
}
