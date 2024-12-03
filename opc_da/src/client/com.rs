use super::Client;

thread_local! {
    pub static COM_INIT: std::cell::RefCell<std::sync::Once> = const { std::cell::RefCell::new(std::sync::Once::new()) };
    pub static COM_RESULT: std::cell::RefCell<windows_core::HRESULT> = const { std::cell::RefCell::new(windows::Win32::Foundation::S_OK) };
}

impl Client {
    pub fn ensure_com() -> windows_core::HRESULT {
        COM_INIT.with(|init| {
            init.borrow().call_once(|| unsafe {
                COM_RESULT.with(|result| {
                    *result.borrow_mut() = windows::Win32::System::Com::CoInitializeEx(
                        None,
                        windows::Win32::System::Com::COINIT_MULTITHREADED,
                    );
                });
            });
        });

        COM_RESULT.with(|result| *result.borrow())
    }
}
