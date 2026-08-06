//! Safe, owning client wrappers for OPC Common interfaces.

use opc_classic_types::{ComObject, Error, ErrorCode, Guid, Result};
use opc_classic_utils::{CoTaskMemArrayOut, CoTaskMemOut, NoCleanup, OwnedPwstr, WideCString};
use windows_core::{GUID, HRESULT, Interface, PCWSTR};

use crate::abi::{
    from_abi_error, from_abi_guid, interface_from_object, interface_from_raw_owned,
    object_from_interface, to_abi_guid,
};
use crate::{IOPCCommon, IOPCEnumGUID, IOPCServerList2, IOPCShutdown};

#[derive(Clone)]
pub struct CommonClient {
    inner: IOPCCommon,
}

impl CommonClient {
    /// Obtains the OPC Common interface from an owning COM object.
    pub fn from_object(object: &ComObject) -> Result<Self> {
        interface_from_object(object).map(Self::from_inner)
    }

    fn from_inner(inner: IOPCCommon) -> Self {
        Self { inner }
    }

    /// Returns an owning, Windows-independent COM identity for this client.
    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn set_locale(&self, locale: u32) -> Result<()> {
        unsafe { self.inner.SetLocaleID(locale) }.map_err(from_abi_error)
    }

    pub fn locale(&self) -> Result<u32> {
        unsafe { self.inner.GetLocaleID() }.map_err(from_abi_error)
    }

    pub fn available_locales(&self) -> Result<Vec<u32>> {
        let mut count = 0u32;
        let mut output = CoTaskMemArrayOut::new_reported(NoCleanup);
        let call = unsafe {
            self.inner
                .QueryAvailableLocaleIDs(&mut count, output.as_mut_ptr())
        };
        call.map_err(from_abi_error)?;
        // The element count is trusted only after the COM call reports success.
        let output = unsafe { output.into_array_with_len(count as usize) }?;
        Ok(output.as_slice().to_vec())
    }

    pub fn error_string(&self, error: ErrorCode) -> Result<String> {
        let mut value = CoTaskMemOut::<u16>::new();
        let call = unsafe {
            (Interface::vtable(&self.inner).GetErrorString)(
                Interface::as_raw(&self.inner),
                HRESULT(error.raw()),
                value.as_mut_ptr().cast(),
            )
        };
        // SAFETY: OPC Common transfers a task-allocated string to the caller.
        // Adopt it before inspecting the status so an abnormal server cannot
        // leak an out value written alongside a failure.
        let value = unsafe { value.into_pwstr() };
        call.ok().map_err(from_abi_error)?;
        Ok(value.to_string_lossy())
    }

    pub fn set_client_name(&self, name: &str) -> Result<()> {
        let name = wide(name)?;
        unsafe { self.inner.SetClientName(PCWSTR(name.as_ptr())) }.map_err(from_abi_error)
    }
}

#[derive(Clone)]
pub struct GuidEnumerator {
    inner: IOPCEnumGUID,
}

impl GuidEnumerator {
    pub fn from_object(object: &ComObject) -> Result<Self> {
        interface_from_object(object).map(Self::from_inner)
    }

    fn from_inner(inner: IOPCEnumGUID) -> Self {
        Self { inner }
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn next_batch(&self, maximum: usize) -> Result<Vec<Guid>> {
        let capacity = u32::try_from(maximum)
            .map_err(|_| Error::invalid_argument("enumeration batch is too large"))?;
        let mut values = vec![GUID::zeroed(); capacity as usize];
        let mut fetched = 0u32;
        unsafe { self.inner.Next(&mut values, &mut fetched) }.map_err(from_abi_error)?;
        if fetched > capacity {
            return Err(Error::unexpected(
                "OPC GUID enumerator returned more values than requested",
            ));
        }
        values.truncate(fetched as usize);
        Ok(values.into_iter().map(from_abi_guid).collect())
    }

    pub fn skip(&self, count: u32) -> Result<()> {
        unsafe { self.inner.Skip(count) }.map_err(from_abi_error)
    }

    pub fn reset(&self) -> Result<()> {
        unsafe { self.inner.Reset() }.map_err(from_abi_error)
    }

    pub fn try_clone(&self) -> Result<Self> {
        let mut raw = core::ptr::null_mut();
        let call = unsafe {
            (Interface::vtable(&self.inner).Clone)(Interface::as_raw(&self.inner), &mut raw)
        };
        let cloned = unsafe { interface_from_raw_owned(raw) };
        call.ok().map_err(from_abi_error)?;
        cloned.map(Self::from_inner)
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
    pub fn from_object(object: &ComObject) -> Result<Self> {
        interface_from_object(object).map(Self::from_inner)
    }

    fn from_inner(inner: IOPCServerList2) -> Self {
        Self { inner }
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn enum_classes(&self, implemented: &[Guid], required: &[Guid]) -> Result<GuidEnumerator> {
        let implemented: Vec<_> = implemented.iter().copied().map(to_abi_guid).collect();
        let required: Vec<_> = required.iter().copied().map(to_abi_guid).collect();
        let implemented_count = u32::try_from(implemented.len())
            .map_err(|_| Error::invalid_argument("too many implemented categories"))?;
        let required_count = u32::try_from(required.len())
            .map_err(|_| Error::invalid_argument("too many required categories"))?;
        let mut raw = core::ptr::null_mut();
        let call = unsafe {
            (Interface::vtable(&self.inner).EnumClassesOfCategories)(
                Interface::as_raw(&self.inner),
                implemented_count,
                implemented.as_ptr(),
                required_count,
                required.as_ptr(),
                &mut raw,
            )
        };
        let enumerator = unsafe { interface_from_raw_owned(raw) };
        call.ok().map_err(from_abi_error)?;
        enumerator.map(GuidEnumerator::from_inner)
    }

    pub fn class_details(&self, class_id: &Guid) -> Result<ClassDetails> {
        let class_id = to_abi_guid(*class_id);
        let mut prog_id = windows_core::PWSTR::null();
        let mut user_type = windows_core::PWSTR::null();
        let mut version_independent = windows_core::PWSTR::null();
        let call = unsafe {
            self.inner.GetClassDetails(
                &class_id,
                &mut prog_id,
                &mut user_type,
                &mut version_independent,
            )
        };
        // Adopt every non-null output before propagating the HRESULT.
        let prog_id = unsafe { OwnedPwstr::from_raw(prog_id.0) };
        let user_type = unsafe { OwnedPwstr::from_raw(user_type.0) };
        let version_independent = unsafe { OwnedPwstr::from_raw(version_independent.0) };
        call.map_err(from_abi_error)?;
        Ok(ClassDetails {
            prog_id: prog_id.to_string_lossy(),
            user_type: user_type.to_string_lossy(),
            version_independent_prog_id: version_independent.to_string_lossy(),
        })
    }

    pub fn class_id_from_prog_id(&self, prog_id: &str) -> Result<Guid> {
        let prog_id = wide(prog_id)?;
        unsafe { self.inner.CLSIDFromProgID(PCWSTR(prog_id.as_ptr())) }
            .map(from_abi_guid)
            .map_err(from_abi_error)
    }
}

#[derive(Clone)]
pub struct ShutdownClient {
    inner: IOPCShutdown,
}

impl ShutdownClient {
    pub fn from_object(object: &ComObject) -> Result<Self> {
        interface_from_object(object).map(Self::from_inner)
    }

    fn from_inner(inner: IOPCShutdown) -> Self {
        Self { inner }
    }

    pub fn object(&self) -> ComObject {
        object_from_interface(&self.inner)
    }

    pub fn request(&self, reason: &str) -> Result<()> {
        let reason = wide(reason)?;
        unsafe { self.inner.ShutdownRequest(PCWSTR(reason.as_ptr())) }.map_err(from_abi_error)
    }
}

fn wide(value: &str) -> Result<WideCString> {
    WideCString::try_from(value)
        .map_err(|_| Error::invalid_argument("string contains an interior NUL"))
}
