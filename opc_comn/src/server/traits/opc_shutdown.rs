pub trait OpcShutdown {
    fn shutdown_request(&self, reason: String) -> windows_core::Result<()>;
}
