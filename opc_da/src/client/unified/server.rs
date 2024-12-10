use crate::client::{v1, v2, v3};

pub enum Server {
    V1(v1::Server),
    V2(v2::Server),
    V3(v3::Server),
}
