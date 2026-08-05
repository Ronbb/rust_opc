use opc_classic_types::{Error, ErrorCode, ErrorKind, Result};

/// OPC Alarms & Events server lifecycle state.
#[repr(i32)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum AeServerState {
    Running = 1,
    Failed = 2,
    NoConfiguration = 3,
    Suspended = 4,
    Test = 5,
    CommunicationFault = 6,
}

impl AeServerState {
    pub(crate) fn from_raw(value: i32) -> Result<Self> {
        match value {
            1 => Ok(Self::Running),
            2 => Ok(Self::Failed),
            3 => Ok(Self::NoConfiguration),
            4 => Ok(Self::Suspended),
            5 => Ok(Self::Test),
            6 => Ok(Self::CommunicationFault),
            _ => Err(Error::new(
                ErrorKind::Protocol,
                ErrorCode::UNEXPECTED,
                "OPC AE server returned an unknown state",
            )),
        }
    }

    pub(crate) const fn raw(self) -> i32 {
        self as i32
    }
}

/// Direction used while navigating the AE area hierarchy.
#[repr(i32)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum BrowseDirection {
    Up = 1,
    Down = 2,
    To = 3,
}

impl BrowseDirection {
    pub(crate) const fn raw(self) -> i32 {
        self as i32
    }
}
