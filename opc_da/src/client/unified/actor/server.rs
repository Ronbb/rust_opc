use actix::prelude::*;

use crate::client::unified::Server;

impl Actor for Server {
    type Context = SyncContext<Self>;
}
