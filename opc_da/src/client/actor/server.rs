use actix::prelude::*;

use crate::client::Server;

impl Actor for Server {
    type Context = Context<Self>;

    fn started(&mut self, ctx: &mut Self::Context) {
        ctx.set_mailbox_capacity(128);
    }
}

pub struct ServerActor(Addr<Server>);

impl ServerActor {
    pub fn new(server: Server) -> Self {
        Self(server.start())
    }
}

// deref to the inner Addr<Server>
impl std::ops::Deref for ServerActor {
    type Target = Addr<Server>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}
