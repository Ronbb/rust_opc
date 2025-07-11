mod client;
mod runtime;
mod server;

#[cfg(test)]
mod tests;

pub use client::*;
pub use runtime::*;

fn mb_error(err: actix::MailboxError) -> windows::core::Error {
    windows::core::Error::new(
        windows::Win32::Foundation::E_FAIL,
        format!("Failed to send message to client actor: {err:?}"),
    )
}

#[macro_export]
macro_rules! mb_error {
    ($err:expr) => {
        $err.map_err($crate::client::unified::actor::mb_error)?
    };
}
