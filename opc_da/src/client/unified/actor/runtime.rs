use crate::client::unified::Guard;

pub fn try_create_runtime() -> std::io::Result<tokio::runtime::Runtime> {
    tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .on_thread_start(Guard::<()>::initialize)
        .on_thread_stop(Guard::<()>::uninitialize)
        .build()
}

pub fn create_runtime() -> tokio::runtime::Runtime {
    try_create_runtime().expect("Failed to create runtime")
}
