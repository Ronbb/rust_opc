use actix::prelude::*;

use crate::{
    client::{unified::Client, GuidIterator},
    mb_error,
    utils::RemotePointer,
};

impl Actor for Client {
    type Context = Context<Self>;
}

impl Actor for GuidIterator {
    type Context = Context<Self>;
}

pub struct ClientActor(Addr<Client>);

impl ClientActor {
    pub fn new(client: Client) -> windows::core::Result<Self> {
        Ok(Self(client.start()))
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
struct GetServerGuids;

impl ClientActor {
    pub async fn get_servers(&self) -> windows::core::Result<Vec<(windows::core::GUID, String)>> {
        mb_error!(self.send(GetServerGuids).await)
    }
}

impl Handler<GetServerGuids> for Client {
    type Result = windows::core::Result<Vec<(windows::core::GUID, String)>>;

    fn handle(&mut self, _: GetServerGuids, _: &mut Self::Context) -> Self::Result {
        self.get_servers()?
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
