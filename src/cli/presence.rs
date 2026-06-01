use clap::{Args, Subcommand};
use std::time::Instant;

use crate::api::presence::PresenceClient;
use crate::api::HttpClient;
use crate::auth::token::TokenSet;
use crate::error::Result;
use crate::output::{self, OutputFormat};

#[derive(Args)]
pub struct PresenceArgs {
    #[command(subcommand)]
    pub command: PresenceCommand,
}

#[derive(Subcommand)]
pub enum PresenceCommand {
    /// Get current presence status
    Get,
    /// Set presence to Available (online)
    Online,
    /// Set presence to Offline
    Offline,
    /// Set presence to Away
    Away,
    /// Set presence to Busy
    Busy,
    /// Set presence to Do Not Disturb
    Dnd,
    /// Set a custom availability status
    Set {
        /// Availability: Available, Busy, DoNotDisturb, BeRightBack, Away, Offline
        #[arg(value_parser = ["Available", "Busy", "DoNotDisturb", "BeRightBack", "Away", "Offline"])]
        availability: String,
    },
}

pub async fn handle(
    args: &PresenceArgs,
    tokens: &TokenSet,
    http: &HttpClient,
    format: OutputFormat,
) -> Result<()> {
    let client = PresenceClient::new(http, tokens);

    match &args.command {
        PresenceCommand::Get => {
            let start = Instant::now();
            let presence = client.get_presence().await?;
            output::print_output(format, presence, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Online => {
            let start = Instant::now();
            client.set_presence("Available").await?;
            let result = serde_json::json!({"availability": "Available", "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Offline => {
            let start = Instant::now();
            client.set_presence("Offline").await?;
            let result = serde_json::json!({"availability": "Offline", "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Away => {
            let start = Instant::now();
            client.set_presence("Away").await?;
            let result = serde_json::json!({"availability": "Away", "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Busy => {
            let start = Instant::now();
            client.set_presence("Busy").await?;
            let result = serde_json::json!({"availability": "Busy", "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Dnd => {
            let start = Instant::now();
            client.set_presence("DoNotDisturb").await?;
            let result = serde_json::json!({"availability": "DoNotDisturb", "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        PresenceCommand::Set { availability } => {
            let start = Instant::now();
            client.set_presence(availability).await?;
            let result = serde_json::json!({"availability": availability, "status": "set"});
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
    }
    Ok(())
}
