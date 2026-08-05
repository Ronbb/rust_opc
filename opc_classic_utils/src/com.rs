use std::marker::PhantomData;
use std::rc::Rc;

use windows::Win32::System::Com::{
    COINIT, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize,
};
use windows_core::Error;

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
    pub fn initialize(model: ApartmentModel) -> Result<Self, Error> {
        let coinit: COINIT = match model {
            ApartmentModel::SingleThreaded => COINIT_APARTMENTTHREADED,
            ApartmentModel::MultiThreaded => COINIT_MULTITHREADED,
        };
        unsafe { CoInitializeEx(None, coinit) }.ok()?;
        Ok(Self {
            model,
            _thread_bound: PhantomData,
        })
    }

    pub fn sta() -> Result<Self, Error> {
        Self::initialize(ApartmentModel::SingleThreaded)
    }

    pub fn mta() -> Result<Self, Error> {
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
