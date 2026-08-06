#![cfg(windows)]

use std::env;
use std::process::Command;

use opc_classic_types::{Error, Result, Value, ValueType};
use opc_classic_utils::ComApartment;
use opc_da_bindings::client::{
    DaClient, DataSource, GroupOptions, ItemSpec, ServerState, WriteValue,
};

const READ_ITEM_ID: &str = "Demo.Ramp";
const WRITE_ITEM_ID: &str = "Simulation.Register_I4";
const READ_CLIENT_HANDLE: u32 = 0x4f50_4301;
const WRITE_CLIENT_HANDLE: u32 = 0x4f50_4302;
const WRITE_VALUE: i32 = 12_345;

#[test]
#[ignore = "requires a locally registered OPC Kit Server"]
fn opc_kit_da_round_trip() {
    let prog_id = env::var("OPC_TEST_PROG_ID")
        .expect("OPC_TEST_PROG_ID must identify the registered OPC DA server");

    if let Err(error) = run_round_trip(&prog_id) {
        panic!(
            "OPC DA integration test failed for ProgID {prog_id}: {error}\n{}",
            diagnostics(&prog_id)
        );
    }
}

fn run_round_trip(prog_id: &str) -> Result<()> {
    let apartment = ComApartment::mta()?;
    let client = DaClient::connect_prog_id(&apartment, prog_id)?;

    let status = client.status()?;
    println!(
        "connected prog_id={prog_id} vendor={:?} version={:?} state={:?}",
        status.vendor_info, status.version, status.state
    );
    ensure(
        status.state == ServerState::Running,
        format!("server state was {:?}, expected Running", status.state),
    )?;
    ensure(
        !status.vendor_info.is_empty(),
        "server returned an empty vendor string".to_owned(),
    )?;

    let group = client.add_group(GroupOptions::new("rust-opc-ci"))?;
    let state = group.state()?;
    println!(
        "created group server_handle={} revised_update_rate={}",
        state.server_handle,
        group.revised_update_rate()
    );
    ensure(state.active, "new group was not active".to_owned())?;
    ensure(
        state.server_handle != 0,
        "server returned a zero group handle".to_owned(),
    )?;

    let added = group.add_items(&[
        ItemSpec::new(READ_ITEM_ID, READ_CLIENT_HANDLE),
        ItemSpec::new(WRITE_ITEM_ID, WRITE_CLIENT_HANDLE),
    ])?;
    ensure(
        added.len() == 2,
        format!("server returned {} item results, expected 2", added.len()),
    )?;
    let mut added = added.into_iter();
    let read_item = added
        .next()
        .ok_or_else(|| Error::unexpected("server returned no read item result"))?
        .map_err(|error| Error::from_code(error.code))?;
    let write_item = added
        .next()
        .ok_or_else(|| Error::unexpected("server returned no write item result"))?
        .map_err(|error| Error::from_code(error.code))?;
    ensure(
        read_item.server_handle.raw() != 0 && write_item.server_handle.raw() != 0,
        "server returned a zero item handle".to_owned(),
    )?;
    ensure(
        read_item.server_handle != write_item.server_handle,
        "server returned duplicate item handles".to_owned(),
    )?;
    ensure(
        read_item.canonical_data_type == ValueType::F64,
        format!(
            "{READ_ITEM_ID} canonical type was {:?}, expected F64",
            read_item.canonical_data_type
        ),
    )?;
    ensure(
        write_item.canonical_data_type == ValueType::I32,
        format!(
            "{WRITE_ITEM_ID} canonical type was {:?}, expected I32",
            write_item.canonical_data_type
        ),
    )?;
    ensure(
        write_item.access_rights & 2 != 0,
        format!(
            "{WRITE_ITEM_ID} is not writable (access rights {})",
            write_item.access_rights
        ),
    )?;

    let timestamped = read_one(&group, read_item.server_handle)?;
    ensure(
        timestamped.client_handle.0 == READ_CLIENT_HANDLE,
        format!(
            "read client handle mismatch: expected {READ_CLIENT_HANDLE}, got {}",
            timestamped.client_handle.0
        ),
    )?;
    ensure(
        timestamped.quality & 0xc0 == 0xc0,
        format!("bad OPC quality 0x{:04x}", timestamped.quality),
    )?;
    ensure(
        timestamped.timestamp.ticks() != 0,
        "server returned an empty timestamp".to_owned(),
    )?;

    let before = read_one(&group, write_item.server_handle)?;
    ensure(
        before.client_handle.0 == WRITE_CLIENT_HANDLE,
        format!(
            "write client handle mismatch: expected {WRITE_CLIENT_HANDLE}, got {}",
            before.client_handle.0
        ),
    )?;

    let original_value = before.value;
    let round_trip = (|| {
        write_one(&group, write_item.server_handle, Value::I32(WRITE_VALUE))?;

        let after = read_one(&group, write_item.server_handle)?;
        ensure(
            after.value == Value::I32(WRITE_VALUE),
            format!("write/read value mismatch: {:?}", after.value),
        )?;
        ensure(
            after.quality & 0xc0 == 0xc0,
            format!("bad OPC quality after write 0x{:04x}", after.quality),
        )?;
        println!(
            "round trip item={WRITE_ITEM_ID} server_handle={} value={:?} quality=0x{:04x} timestamp={}",
            write_item.server_handle.raw(),
            after.value,
            after.quality,
            after.timestamp.ticks()
        );
        Ok(())
    })();
    let restore =
        write_one(&group, write_item.server_handle, original_value.clone()).and_then(|()| {
            let restored = read_one(&group, write_item.server_handle)?;
            ensure(
                restored.value == original_value,
                format!("original item value was not restored: {:?}", restored.value),
            )
        });
    match (round_trip, restore) {
        (Ok(()), Ok(())) => {}
        (Err(error), Ok(())) => return Err(error),
        (Ok(()), Err(restore_error)) => {
            let message = format!("failed to restore original item value: {restore_error}");
            return Err(restore_error.with_message(message));
        }
        (Err(error), Err(restore_error)) => {
            return Err(Error::unexpected(format!(
                "round trip failed: {error}; restoring original item value also failed: {restore_error}"
            )));
        }
    }
    group.close(true)?;
    Ok(())
}

fn write_one(
    group: &opc_da_bindings::client::DaGroup<'_>,
    handle: opc_da_bindings::client::ServerItemHandle,
    value: Value,
) -> Result<()> {
    group
        .write(&[WriteValue {
            server_handle: handle,
            value,
        }])?
        .into_iter()
        .next()
        .ok_or_else(|| Error::unexpected("server returned no write result"))?
        .map_err(|error| Error::from_code(error.code))
}

fn read_one(
    group: &opc_da_bindings::client::DaGroup<'_>,
    handle: opc_da_bindings::client::ServerItemHandle,
) -> Result<opc_da_bindings::client::Sample> {
    group
        .read(DataSource::Device, &[handle])?
        .into_iter()
        .next()
        .ok_or_else(|| Error::unexpected("server returned no read result"))?
        .map_err(|error| Error::from_code(error.code))
}

fn ensure(condition: bool, message: String) -> Result<()> {
    if condition {
        Ok(())
    } else {
        Err(Error::unexpected(message))
    }
}

fn diagnostics(prog_id: &str) -> String {
    format!(
        "diagnostics:\n  prog_id={prog_id}\n  tasklist={}\n  registry32={}\n  registry64={}\n  where_opcrtkit={}",
        command_output("tasklist", &["/FI", "IMAGENAME eq opcrtkit.exe"]),
        command_output(
            "reg",
            &["query", &format!("HKCR\\{prog_id}"), "/s", "/reg:32"],
        ),
        command_output(
            "reg",
            &["query", &format!("HKCR\\{prog_id}"), "/s", "/reg:64"],
        ),
        command_output("where", &["opcrtkit"]),
    )
}

fn command_output(program: &str, args: &[&str]) -> String {
    match Command::new(program).args(args).output() {
        Ok(output) => format!(
            "status={} stdout={} stderr={}",
            output.status,
            String::from_utf8_lossy(&output.stdout).trim(),
            String::from_utf8_lossy(&output.stderr).trim()
        ),
        Err(error) => format!("spawn failed: {error}"),
    }
}
