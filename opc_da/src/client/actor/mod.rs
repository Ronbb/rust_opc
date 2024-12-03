mod client;
mod server;

pub use client::*;
pub use server::*;

fn convert_error(err: actix::MailboxError) -> windows_core::Error {
    windows_core::Error::new(
        windows::Win32::Foundation::E_FAIL,
        format!("Failed to send message to client actor: {:?}", err),
    )
}

#[macro_export]
macro_rules! convert_error {
    ($err:expr) => {
        $err.map_err($crate::client::actor::convert_error)?
    };
}
