use actix::prelude::*;

use crate::{
    client::{Client, ServerFilter},
    convert_error,
};

impl Actor for Client {
    type Context = Context<Self>;

    fn started(&mut self, ctx: &mut Self::Context) {
        ctx.set_mailbox_capacity(128);
    }
}

pub struct ClientActor(Addr<Client>);

impl ClientActor {
    pub fn new() -> windows_core::Result<Self> {
        Ok(Self(Client::new()?.start()))
    }
}

// deref to the inner Addr<Client>
impl std::ops::Deref for ClientActor {
    type Target = Addr<Client>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

#[derive(Message)]
#[rtype(result = "windows_core::Result<Vec<(windows_core::GUID, String)>>")]
struct GetServerGuids(pub ServerFilter);

impl ClientActor {
    pub async fn get_servers(
        &self,
        filter: ServerFilter,
    ) -> windows_core::Result<Vec<(windows_core::GUID, String)>> {
        convert_error!(self.send(GetServerGuids(filter)).await)
    }
}

impl Handler<GetServerGuids> for Client {
    type Result = windows_core::Result<Vec<(windows_core::GUID, String)>>;

    fn handle(&mut self, message: GetServerGuids, _: &mut Self::Context) -> Self::Result {
        self.get_servers(message.0)?
            .map(|r| match r {
                Ok(guid) => {
                    let name = unsafe {
                        windows::Win32::System::Com::ProgIDFromCLSID(&guid)
                            .map_err(|e| windows_core::Error::new(e.code(), "Failed to get ProgID"))
                    }?;

                    Ok((guid, unsafe { name.to_string() }?))
                }
                Err(e) => Err(e),
            })
            .collect()
    }
}
