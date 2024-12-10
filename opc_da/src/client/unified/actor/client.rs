use actix::prelude::*;

use crate::{
    client::{unified::Client, RemotePointer},
    convert_error, def,
};

impl Actor for Client {
    type Context = Context<Self>;

    fn started(&mut self, ctx: &mut Self::Context) {
        ctx.set_mailbox_capacity(128);
    }
}

pub struct ClientActor(Addr<Client>);

impl ClientActor {
    pub fn new() -> windows::core::Result<Self> {
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
#[rtype(result = "windows::core::Result<Vec<(windows::core::GUID, String)>>")]
struct GetServerGuids(pub def::ServerFilter);

impl ClientActor {
    pub async fn get_servers(
        &self,
        filter: def::ServerFilter,
    ) -> windows::core::Result<Vec<(windows::core::GUID, String)>> {
        convert_error!(self.send(GetServerGuids(filter)).await)
    }
}

impl Handler<GetServerGuids> for Client {
    type Result = windows::core::Result<Vec<(windows::core::GUID, String)>>;

    fn handle(&mut self, message: GetServerGuids, _: &mut Self::Context) -> Self::Result {
        self.get_servers(message.0)?
            .map(|r| match r {
                Ok(guid) => {
                    let name = unsafe {
                        windows::Win32::System::Com::ProgIDFromCLSID(&guid).map_err(|e| {
                            windows::core::Error::new(e.code(), "Failed to get ProgID")
                        })
                    }?;

                    let name = RemotePointer::from(name);

                    Ok((guid, name.try_into()?))
                }
                Err(e) => Err(e),
            })
            .collect()
    }
}