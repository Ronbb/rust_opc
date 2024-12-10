use windows::core::Interface;

use crate::client::{
    traits::{AsyncIo2Trait, DataObjectTrait, GroupStateMgtTrait, ItemMgtTrait},
    AsyncIo3Trait, AsyncIoTrait, ConnectionPointContainerTrait, GroupStateMgt2Trait,
    ItemDeadbandMgtTrait, ItemSamplingMgtTrait, PublicGroupStateMgtTrait, SyncIo2Trait,
    SyncIoTrait,
};

/*
opc_da_bindings::IOPCItemMgt,
opc_da_bindings::IOPCGroupStateMgt,
opc_da_bindings::IOPCGroupStateMgt2,
opc_da_bindings::IOPCPublicGroupStateMgt,
opc_da_bindings::IOPCSyncIO,
opc_da_bindings::IOPCSyncIO2,
opc_da_bindings::IOPCAsyncIO2,
opc_da_bindings::IOPCAsyncIO3,
opc_da_bindings::IOPCItemDeadbandMgt,
opc_da_bindings::IOPCItemSamplingMgt,
windows::Win32::System::Com::IConnectionPointContainer,
opc_da_bindings::IOPCAsyncIO,
windows::Win32::System::Com::IDataObject
*/
pub struct Group {
    // 1.0 required
    // 2.0 required
    // 3.0 required
    pub(crate) item_mgt: opc_da_bindings::IOPCItemMgt,
    // 1.0 required
    // 2.0 required
    // 3.0 required
    pub(crate) group_state_mgt: opc_da_bindings::IOPCGroupStateMgt,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) group_state_mgt2: Option<opc_da_bindings::IOPCGroupStateMgt2>,
    // 1.0 optional
    // 2.0 optional
    // 3.0 N/A
    pub(crate) public_group_state_mgt: Option<opc_da_bindings::IOPCPublicGroupStateMgt>,
    // 1.0 required
    // 2.0 required
    // 3.0 required
    pub(crate) sync_io: opc_da_bindings::IOPCSyncIO,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) sync_io2: Option<opc_da_bindings::IOPCSyncIO2>,
    // 1.0 required
    // 2.0 optional
    // 3.0 N/A
    pub(crate) async_io: Option<opc_da_bindings::IOPCAsyncIO>,
    // 1.0 N/A
    // 2.0 required
    // 3.0 required
    pub(crate) async_io2: Option<opc_da_bindings::IOPCAsyncIO2>,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) async_io3: Option<opc_da_bindings::IOPCAsyncIO3>,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 required
    pub(crate) item_deadband_mgt: Option<opc_da_bindings::IOPCItemDeadbandMgt>,
    // 1.0 N/A
    // 2.0 N/A
    // 3.0 optional
    pub(crate) item_sampling_mgt: Option<opc_da_bindings::IOPCItemSamplingMgt>,
    // 1.0 N/A
    // 2.0 required
    // 3.0 required
    pub(crate) connection_point_container:
        Option<windows::Win32::System::Com::IConnectionPointContainer>,
    // 1.0 required
    // 2.0 optional
    // 3.0 N/A
    pub(crate) data_object: Option<windows::Win32::System::Com::IDataObject>,
}

impl TryFrom<windows::core::IUnknown> for Group {
    type Error = windows::core::Error;

    fn try_from(value: windows::core::IUnknown) -> windows::core::Result<Self> {
        let item_mgt = value.cast::<opc_da_bindings::IOPCItemMgt>()?;
        let group_state_mgt = value.cast::<opc_da_bindings::IOPCGroupStateMgt>()?;
        let group_state_mgt2 = value.cast::<opc_da_bindings::IOPCGroupStateMgt2>().ok();
        let public_group_state_mgt = value
            .cast::<opc_da_bindings::IOPCPublicGroupStateMgt>()
            .ok();
        let sync_io = value.cast::<opc_da_bindings::IOPCSyncIO>()?;
        let sync_io2 = value.cast::<opc_da_bindings::IOPCSyncIO2>().ok();
        let async_io = value.cast::<opc_da_bindings::IOPCAsyncIO>().ok();
        let async_io2 = value.cast::<opc_da_bindings::IOPCAsyncIO2>().ok();
        let async_io3 = value.cast::<opc_da_bindings::IOPCAsyncIO3>().ok();
        let item_deadband_mgt = value.cast::<opc_da_bindings::IOPCItemDeadbandMgt>().ok();
        let item_sampling_mgt = value.cast::<opc_da_bindings::IOPCItemSamplingMgt>().ok();
        let connection_point_container = value
            .cast::<windows::Win32::System::Com::IConnectionPointContainer>()
            .ok();
        let data_object = value
            .cast::<windows::Win32::System::Com::IDataObject>()
            .ok();

        Ok(Self {
            item_mgt,
            group_state_mgt,
            group_state_mgt2,
            public_group_state_mgt,
            sync_io,
            sync_io2,
            async_io,
            async_io2,
            async_io3,
            item_deadband_mgt,
            item_sampling_mgt,
            connection_point_container,
            data_object,
        })
    }
}

impl ItemMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemMgt> {
        Ok(&self.item_mgt)
    }
}

impl GroupStateMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt> {
        Ok(&self.group_state_mgt)
    }
}

impl GroupStateMgt2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCGroupStateMgt2> {
        self.group_state_mgt2.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCGroupStateMgt2 not supported",
            )
        })
    }
}

impl PublicGroupStateMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCPublicGroupStateMgt> {
        self.public_group_state_mgt.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCPublicGroupStateMgt not supported",
            )
        })
    }
}

impl SyncIoTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO> {
        Ok(&self.sync_io)
    }
}

impl SyncIo2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCSyncIO2> {
        self.sync_io2.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCSyncIO2 not supported",
            )
        })
    }
}

impl AsyncIoTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO> {
        self.async_io.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCAsyncIO not supported",
            )
        })
    }
}

impl AsyncIo2Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO2> {
        self.async_io2.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCAsyncIO2 not supported",
            )
        })
    }
}

impl AsyncIo3Trait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCAsyncIO3> {
        self.async_io3.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCAsyncIO3 not supported",
            )
        })
    }
}

impl ItemDeadbandMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemDeadbandMgt> {
        self.item_deadband_mgt.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCItemDeadbandMgt not supported",
            )
        })
    }
}

impl ItemSamplingMgtTrait for Group {
    fn interface(&self) -> windows::core::Result<&opc_da_bindings::IOPCItemSamplingMgt> {
        self.item_sampling_mgt.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IOPCItemSamplingMgt not supported",
            )
        })
    }
}

impl ConnectionPointContainerTrait for Group {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer> {
        self.connection_point_container.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IConnectionPointContainer not supported",
            )
        })
    }
}

impl DataObjectTrait for Group {
    fn interface(&self) -> windows::core::Result<&windows::Win32::System::Com::IDataObject> {
        self.data_object.as_ref().ok_or_else(|| {
            windows::core::Error::new(
                windows::Win32::Foundation::E_NOTIMPL,
                "IDataObject not supported",
            )
        })
    }
}