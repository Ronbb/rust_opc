mod opc_common;

pub use opc_common::{
    OPCCommon, OPCEnumGUID, OPCServerList, OPCServerList2, OPCShutdown,
    traits::{
        OPCCommon as OPCCommonTrait, OPCEnumGUID as OPCEnumGUIDTrait,
        OPCServerList as OPCServerListTrait, OPCServerList2 as OPCServerList2Trait,
        OPCShutdown as OPCShutdownTrait,
    },
};
