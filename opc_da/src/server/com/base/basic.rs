use super::variant::Variant;

#[derive(Clone, Default)]
pub struct Quality(pub u16);

#[derive(Default)]
pub struct Value {
    pub variant: Variant,
    pub quality: Quality,
    pub timestamp: Option<std::time::SystemTime>,
}

#[derive(Default)]
pub struct AccessRight {
    pub readable: bool,
    pub writable: bool,
}
