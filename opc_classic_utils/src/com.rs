use std::marker::PhantomData;
use std::rc::Rc;

use opc_classic_types::{Error, ErrorCode, Result};
use windows::Win32::System::Com::{
    COINIT, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ApartmentModel {
    SingleThreaded,
    MultiThreaded,
}

/// Balances COM initialization on the current thread.
///
/// The `Rc` marker deliberately makes the token `!Send` and `!Sync`, ensuring
/// that `CoUninitialize` runs on the thread that initialized COM.
pub struct ComApartment {
    model: ApartmentModel,
    _thread_bound: PhantomData<Rc<()>>,
}

impl ComApartment {
    pub fn initialize(model: ApartmentModel) -> Result<Self> {
        let coinit: COINIT = match model {
            ApartmentModel::SingleThreaded => COINIT_APARTMENTTHREADED,
            ApartmentModel::MultiThreaded => COINIT_MULTITHREADED,
        };
        unsafe { CoInitializeEx(None, coinit) }
            .ok()
            .map_err(|error| {
                Error::from_code(ErrorCode::from_raw(error.code().0)).with_message(error.message())
            })?;
        Ok(Self {
            model,
            _thread_bound: PhantomData,
        })
    }

    pub fn sta() -> Result<Self> {
        Self::initialize(ApartmentModel::SingleThreaded)
    }

    pub fn mta() -> Result<Self> {
        Self::initialize(ApartmentModel::MultiThreaded)
    }

    pub fn model(&self) -> ApartmentModel {
        self.model
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}
