use windows::Win32::System::Com::CoRegisterClassObject;
use windows_core::GUID;

pub struct Builder {}

impl Builder {
    pub fn register() {
        let id = GUID::from("c8c381a2-acca-4b22-b867-ebfb49e5beaa");
        
    }
}
