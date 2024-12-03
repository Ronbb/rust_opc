use actix::prelude::*;

use super::{Client, ServerFilter};

impl Actor for Client {
    type Context = Context<Self>;

    fn started(&mut self, ctx: &mut Self::Context) {
        ctx.set_mailbox_capacity(128);
    }
}

pub struct ClientAsync(Addr<Client>);

impl From<Client> for ClientAsync {
    fn from(value: Client) -> Self {
        Self(Actor::start(value))
    }
}

// deref to the inner Addr<Client>
impl std::ops::Deref for ClientAsync {
    type Target = Addr<Client>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

fn convert_error(err: MailboxError) -> windows_core::Error {
    windows_core::Error::new(
        windows::Win32::Foundation::E_FAIL,
        format!("Failed to send message to client actor: {:?}", err),
    )
}

macro_rules! convert_error {
    ($err:expr) => {
        $err.map_err(convert_error)?
    };
}

#[derive(Message)]
#[rtype(result = "windows_core::Result<Vec<windows_core::GUID>>")]
struct GetServerGuids(pub ServerFilter);

impl ClientAsync {
    pub async fn get_servers(
        &self,
        filter: ServerFilter,
    ) -> windows_core::Result<Vec<windows_core::GUID>> {
        convert_error!(self.send(GetServerGuids(filter)).await)
    }
}

impl Handler<GetServerGuids> for Client {
    type Result = windows_core::Result<Vec<windows_core::GUID>>;

    fn handle(&mut self, message: GetServerGuids, _: &mut Self::Context) -> Self::Result {
        self.get_servers(message.0)?.collect()
    }
}
