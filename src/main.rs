use clap::{Parser, Subcommand, ValueEnum};
use serde::Serialize;
use crossterm::{
    event::{self, Event, KeyCode, KeyEventKind},
    execute,
    terminal::{EnterAlternateScreen, LeaveAlternateScreen, disable_raw_mode, enable_raw_mode},
};
use outlook_pst::{
    ltp::prop_context::PropertyValue,
    messaging::{folder::UnicodeFolder, message::UnicodeMessage, store::UnicodeStore},
    ndb::node_id::NodeId,
    *,
};
use ratatui::{
    Terminal,
    backend::CrosstermBackend,
    layout::{Constraint, Layout, Rect},
    style::{Modifier, Style},
    text::{Line, Span, Text},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table, TableState},
};
use chrono::{TimeZone, Utc};
use rusqlite::{Connection, params};
use std::{io, path::PathBuf, rc::Rc, time::Instant};

/// Convert a Windows FILETIME (100-ns ticks since 1601-01-01 UTC) to a
/// human-readable UTC string, e.g. "2010-11-24 15:24:27 UTC".
/// Decompress PR_RTF_COMPRESSED bytes and extract plain text.
fn rtf_compressed_to_text(data: &[u8]) -> Option<String> {
    let rtf = compressed_rtf::decompress_rtf(data).ok()?;
    let doc = rtf_parser::RtfDocument::try_from(rtf.as_str()).ok()?;
    Some(doc.get_text())
}

/// Strip HTML tags and decode common HTML entities for plain-text display.
fn html_to_text(html: &str) -> String {
    let mut out = String::with_capacity(html.len());
    let mut in_tag = false;
    let mut in_style = false;
    let mut in_script = false;
    let mut tag_buf = String::new();

    let mut chars = html.chars().peekable();
    while let Some(c) = chars.next() {
        if in_tag {
            if c == '>' {
                let tag_lower = tag_buf.trim().to_ascii_lowercase();
                if tag_lower.starts_with("style") {
                    in_style = true;
                } else if tag_lower.starts_with("/style") {
                    in_style = false;
                } else if tag_lower.starts_with("script") {
                    in_script = true;
                } else if tag_lower.starts_with("/script") {
                    in_script = false;
                } else if !in_style && !in_script {
                    // Block-level tags → newline
                    let t = tag_lower.split_whitespace().next().unwrap_or("");
                    if matches!(t, "br" | "br/" | "p" | "/p" | "div" | "/div"
                        | "tr" | "/tr" | "li" | "/li" | "h1" | "h2" | "h3"
                        | "h4" | "h5" | "h6" | "/h1" | "/h2" | "/h3"
                        | "/h4" | "/h5" | "/h6") {
                        out.push('\n');
                    }
                }
                tag_buf.clear();
                in_tag = false;
            } else {
                tag_buf.push(c);
            }
        } else if c == '<' {
            in_tag = true;
            tag_buf.clear();
        } else if !in_style && !in_script {
            if c == '&' {
                // Collect entity
                let mut entity = String::new();
                for ec in chars.by_ref() {
                    if ec == ';' { break; }
                    entity.push(ec);
                }
                match entity.as_str() {
                    "amp"  => out.push('&'),
                    "lt"   => out.push('<'),
                    "gt"   => out.push('>'),
                    "quot" => out.push('"'),
                    "apos" => out.push('\''),
                    "nbsp" => out.push(' '),
                    s if s.starts_with('#') => {
                        let n: Option<u32> = if s.starts_with("#x") || s.starts_with("#X") {
                            u32::from_str_radix(&s[2..], 16).ok()
                        } else {
                            s[1..].parse().ok()
                        };
                        if let Some(n) = n.and_then(char::from_u32) {
                            out.push(n);
                        }
                    }
                    _ => { out.push('&'); out.push_str(&entity); out.push(';'); }
                }
            } else if c == '\n' || c == '\r' || c == '\t' {
                // Per HTML spec, whitespace in text nodes collapses to a single
                // space. Only block-level tags (br, p, div…) produce newlines.
                let last = out.chars().next_back();
                if last.map_or(false, |ch| ch != ' ' && ch != '\n') {
                    out.push(' ');
                }
            } else {
                out.push(c);
            }
        }
    }

    // Collapse runs of blank lines to at most one blank line
    let mut result = String::with_capacity(out.len());
    let mut blank_lines = 0u32;
    for line in out.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            blank_lines += 1;
            if blank_lines <= 1 {
                result.push('\n');
            }
        } else {
            blank_lines = 0;
            result.push_str(trimmed);
            result.push('\n');
        }
    }
    result
}

fn filetime_to_string(ticks: i64) -> String {
    // Seconds between 1601-01-01 and 1970-01-01
    const EPOCH_DIFF_SECS: i64 = 11_644_473_600;
    let secs = ticks / 10_000_000 - EPOCH_DIFF_SECS;
    let nanos = ((ticks % 10_000_000) * 100) as u32;
    match Utc.timestamp_opt(secs, nanos) {
        chrono::LocalResult::Single(dt) => dt.format("%Y-%m-%d %H:%M:%S UTC").to_string(),
        _ => format!("(invalid: {})", ticks),
    }
}

#[derive(Clone, ValueEnum)]
enum OutputFormat {
    Csv,
    Tsv,
    Json,
}

#[derive(Parser)]
#[command(name = "pstexplorer")]
#[command(about = "A CLI tool to explore PST files", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// List all emails in a PST file
    List {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
        /// Output format (csv, tsv, or json). Omit for human-readable output.
        #[arg(long)]
        format: Option<OutputFormat>,
        /// Maximum number of entries to output (0 = no limit)
        #[arg(long, default_value_t = 0)]
        limit: usize,
    },
    /// Search emails in a PST file by query string (matches from, to, cc, body)
    Search {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
        /// Search query (case-insensitive, matched against from, to, cc, and body)
        #[arg(required = true)]
        query: String,
        /// Output format (csv, tsv, or json). Omit for human-readable output.
        #[arg(long)]
        format: Option<OutputFormat>,
    },
    /// Browse PST file contents in a TUI
    Browse {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
        /// Enable debug/diagnostic mode: logs events to pstexplorer-debug.log
        #[arg(long)]
        debug: bool,
    },
    /// Print statistics about a PST file
    Stats {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
    },
    /// Export a PST file to a SQLite database
    Export {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
        /// Path for the output SQLite database (default: <pst-name>.db)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },
}

/// Accumulated statistics gathered while walking the PST folder tree.
struct PstStats {
    folder_count: usize,
    email_count: usize,
    attachment_count: usize,
    calendar_count: usize,
    contact_count: usize,
    task_count: usize,
    note_count: usize,
    earliest_ts: Option<i64>,
    latest_ts: Option<i64>,
}

impl PstStats {
    fn new() -> Self {
        PstStats {
            folder_count: 0,
            email_count: 0,
            attachment_count: 0,
            calendar_count: 0,
            contact_count: 0,
            task_count: 0,
            note_count: 0,
            earliest_ts: None,
            latest_ts: None,
        }
    }

    fn update_timestamp(&mut self, ts: i64) {
        self.earliest_ts = Some(match self.earliest_ts {
            Some(e) => e.min(ts),
            None => ts,
        });
        self.latest_ts = Some(match self.latest_ts {
            Some(l) => l.max(ts),
            None => ts,
        });
    }
}

fn collect_stats(
    store: Rc<UnicodeStore>,
    folder: &UnicodeFolder,
    stats: &mut PstStats,
) {
    stats.folder_count += 1;

    if let Some(contents_table) = folder.contents_table() {
        for row in contents_table.rows_matrix() {
            let row_id = u32::from(row.id());
            let entry_id = match store.properties().make_entry_id(NodeId::from(row_id)) {
                Ok(id) => id,
                Err(_) => continue,
            };

            let message = match UnicodeMessage::read(
                Rc::clone(&store),
                &entry_id,
                // Subject, MessageClass, ReceivedTime, ClientSubmitTime, AttachCount
                Some(&[0x0037, 0x001A, 0x0E06, 0x0039, 0x0E13]),
            ) {
                Ok(m) => m,
                Err(_) => continue,
            };

            let props = message.properties();

            // Determine item type from PR_MESSAGE_CLASS (0x001A)
            let message_class: String = props
                .get(0x001A)
                .and_then(|v| match v {
                    PropertyValue::String8(s) => Some(s.to_string().to_ascii_uppercase()),
                    PropertyValue::Unicode(s) => Some(s.to_string().to_ascii_uppercase()),
                    _ => None,
                })
                .unwrap_or_default();

            if message_class.starts_with("IPM.NOTE") || message_class.is_empty() || message_class == "IPM" {
                stats.email_count += 1;
            } else if message_class.starts_with("IPM.APPOINTMENT") || message_class.starts_with("IPM.SCHEDULE") {
                stats.calendar_count += 1;
            } else if message_class.starts_with("IPM.CONTACT") {
                stats.contact_count += 1;
            } else if message_class.starts_with("IPM.TASK") {
                stats.task_count += 1;
            } else if message_class.starts_with("IPM.STICKYNOTE") {
                stats.note_count += 1;
            } else {
                // Treat anything else as an email-like item
                stats.email_count += 1;
            }

            // PR_ATTACH_NUM (0x0E13) gives the count of attachments on this message
            if let Some(PropertyValue::Integer32(n)) = props.get(0x0E13) {
                if *n > 0 {
                    stats.attachment_count += *n as usize;
                }
            }

            // Record timestamp: prefer PR_CLIENT_SUBMIT_TIME (0x0039), fall back to
            // PR_MESSAGE_DELIVERY_TIME (0x0E06)
            let ts = props
                .get(0x0039)
                .or_else(|| props.get(0x0E06))
                .and_then(|v| match v {
                    PropertyValue::Time(t) => Some(*t),
                    _ => None,
                });
            if let Some(t) = ts {
                stats.update_timestamp(t);
            }
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for row in hierarchy_table.rows_matrix() {
            let Ok(entry_id) = store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
            else {
                continue;
            };
            let Ok(subfolder) = UnicodeFolder::read(Rc::clone(&store), &entry_id) else {
                continue;
            };
            collect_stats(Rc::clone(&store), &subfolder, stats);
        }
    }
}

fn stats_pst(file_path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = UnicodeFolder::read(Rc::clone(&store), &ipm_sub_tree_entry_id)?;

    let mut stats = PstStats::new();
    collect_stats(Rc::clone(&store), &ipm_subtree_folder, &mut stats);

    let total_items = stats.email_count
        + stats.calendar_count
        + stats.contact_count
        + stats.task_count
        + stats.note_count;

    println!("PST Statistics: {:?}", file_path);
    println!("  Folders:          {}", stats.folder_count);
    println!("  Total items:      {}", total_items);
    println!("  Emails:           {}", stats.email_count);
    if stats.calendar_count > 0 {
        println!("  Calendar items:   {}", stats.calendar_count);
    }
    if stats.contact_count > 0 {
        println!("  Contacts:         {}", stats.contact_count);
    }
    if stats.task_count > 0 {
        println!("  Tasks:            {}", stats.task_count);
    }
    if stats.note_count > 0 {
        println!("  Notes:            {}", stats.note_count);
    }
    println!("  Attachments:      {}", stats.attachment_count);
    match (stats.earliest_ts, stats.latest_ts) {
        (Some(e), Some(l)) => {
            println!("  Earliest message: {}", filetime_to_string(e));
            println!("  Latest message:   {}", filetime_to_string(l));
        }
        _ => {
            println!("  Date range:       (no timestamps found)");
        }
    }
    Ok(())
}

/// A single email record collected during folder traversal.
#[derive(Serialize)]
struct EmailRecord {
    folder: String,
    subject: String,
    from: String,
    date: String,
}

fn list_emails(
    file_path: &PathBuf,
    format: Option<&OutputFormat>,
    limit: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    // Open the PST file
    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);

    // Get the IPM subtree (where emails are stored)
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = outlook_pst::messaging::folder::UnicodeFolder::read(
        Rc::clone(&store),
        &ipm_sub_tree_entry_id,
    )?;

    match format {
        Some(fmt) => {
            // Collect all emails then output in the requested format
            let mut records: Vec<EmailRecord> = Vec::new();
            collect_emails(Rc::clone(&store), &ipm_subtree_folder, &mut records)?;

            // Apply limit if set
            let records: Vec<EmailRecord> = if limit > 0 {
                records.into_iter().take(limit).collect()
            } else {
                records
            };

            match fmt {
                OutputFormat::Json => {
                    println!("{}", serde_json::to_string_pretty(&records)?);
                }
                OutputFormat::Csv => {
                    println!("folder,subject,from,date");
                    for r in &records {
                        println!(
                            "{},{},{},{}",
                            csv_escape(&r.folder),
                            csv_escape(&r.subject),
                            csv_escape(&r.from),
                            csv_escape(&r.date),
                        );
                    }
                }
                OutputFormat::Tsv => {
                    println!("folder\tsubject\tfrom\tdate");
                    for r in &records {
                        println!(
                            "{}\t{}\t{}\t{}",
                            tsv_escape(&r.folder),
                            tsv_escape(&r.subject),
                            tsv_escape(&r.from),
                            tsv_escape(&r.date),
                        );
                    }
                }
            }
        }
        None => {
            // Original human-readable output
            println!("Listing emails from: {:?}", file_path);

            let hierarchy_table = store.root_hierarchy_table()?;
            println!("PST File Information:");
            println!(
                "  Number of rows in hierarchy: {}",
                hierarchy_table.rows_matrix().count()
            );

            let mut printed = 0usize;
            let total_emails =
                traverse_folder_hierarchy(Rc::clone(&store), &ipm_subtree_folder, limit, &mut printed)?;
            println!("\nFound {} emails in the PST file", total_emails);
        }
    }

    Ok(())
}

/// Escape a field for CSV output (RFC 4180).
fn csv_escape(field: &str) -> String {
    if field.contains(',') || field.contains('"') || field.contains('\n') {
        format!("\"{}\"", field.replace('"', "\"\""))
    } else {
        field.to_string()
    }
}

/// Escape a field for TSV output (replace tabs and newlines with spaces).
fn tsv_escape(field: &str) -> String {
    field.replace('\t', " ").replace('\n', " ")
}

/// Recursively collect email records from the folder tree.
fn collect_emails(
    store: Rc<UnicodeStore>,
    folder: &outlook_pst::messaging::folder::UnicodeFolder,
    records: &mut Vec<EmailRecord>,
) -> Result<(), Box<dyn std::error::Error>> {
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    if let Some(contents_table) = folder.contents_table() {
        for message_row in contents_table.rows_matrix() {
            let message_entry_id =
                store
                    .properties()
                    .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                        message_row.id(),
                    )))?;

            if let Ok(message) = outlook_pst::messaging::message::UnicodeMessage::read(
                store.clone(),
                &message_entry_id,
                Some(&[0x0037, 0x0C1A, 0x0E06]),
            ) {
                let properties = message.properties();

                let subject = properties
                    .get(0x0037)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::String8(s) => {
                            Some(s.to_string())
                        }
                        outlook_pst::ltp::prop_context::PropertyValue::Unicode(s) => {
                            Some(s.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "No Subject".to_string());

                let sender = properties
                    .get(0x0C1A)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::String8(s) => {
                            Some(s.to_string())
                        }
                        outlook_pst::ltp::prop_context::PropertyValue::Unicode(s) => {
                            Some(s.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "Unknown Sender".to_string());

                let received_time = properties
                    .get(0x0E06)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::Time(t) => {
                            Some(t.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "Unknown Date".to_string());

                records.push(EmailRecord {
                    folder: folder_name.clone(),
                    subject,
                    from: sender,
                    date: received_time,
                });
            }
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for subfolder_row in hierarchy_table.rows_matrix() {
            let subfolder_entry_id =
                store
                    .properties()
                    .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                        subfolder_row.id(),
                    )))?;
            let subfolder = outlook_pst::messaging::folder::UnicodeFolder::read(
                store.clone(),
                &subfolder_entry_id,
            )?;
            collect_emails(store.clone(), &subfolder, records)?;
        }
    }

    Ok(())
}

/// A single search result record for structured output.
#[derive(Serialize)]
struct SearchRecord {
    folder: String,
    subject: String,
    from: String,
    to: String,
    cc: String,
    date: String,
}

fn search_emails(
    file_path: &PathBuf,
    query: &str,
    format: Option<&OutputFormat>,
) -> Result<(), Box<dyn std::error::Error>> {
    let query_lower = query.to_ascii_lowercase();
    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = outlook_pst::messaging::folder::UnicodeFolder::read(
        Rc::clone(&store),
        &ipm_sub_tree_entry_id,
    )?;

    match format {
        Some(fmt) => {
            let mut records: Vec<SearchRecord> = Vec::new();
            collect_search_matches(
                Rc::clone(&store),
                &ipm_subtree_folder,
                &query_lower,
                &mut records,
            )?;

            match fmt {
                OutputFormat::Json => {
                    println!("{}", serde_json::to_string_pretty(&records)?);
                }
                OutputFormat::Csv => {
                    println!("folder,subject,from,to,cc,date");
                    for r in &records {
                        println!(
                            "{},{},{},{},{},{}",
                            csv_escape(&r.folder),
                            csv_escape(&r.subject),
                            csv_escape(&r.from),
                            csv_escape(&r.to),
                            csv_escape(&r.cc),
                            csv_escape(&r.date),
                        );
                    }
                }
                OutputFormat::Tsv => {
                    println!("folder\tsubject\tfrom\tto\tcc\tdate");
                    for r in &records {
                        println!(
                            "{}\t{}\t{}\t{}\t{}\t{}",
                            tsv_escape(&r.folder),
                            tsv_escape(&r.subject),
                            tsv_escape(&r.from),
                            tsv_escape(&r.to),
                            tsv_escape(&r.cc),
                            tsv_escape(&r.date),
                        );
                    }
                }
            }
        }
        None => {
            let total = search_traverse_folders(
                Rc::clone(&store),
                &ipm_subtree_folder,
                &query_lower,
            )?;
            println!("\nFound {} matching emails", total);
        }
    }

    Ok(())
}

/// Recursively collect matching search results into a Vec.
fn collect_search_matches(
    store: Rc<UnicodeStore>,
    folder: &outlook_pst::messaging::folder::UnicodeFolder,
    query_lower: &str,
    records: &mut Vec<SearchRecord>,
) -> Result<(), Box<dyn std::error::Error>> {
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    if let Some(contents_table) = folder.contents_table() {
        for message_row in contents_table.rows_matrix() {
            let message_entry_id = store
                .properties()
                .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                    message_row.id(),
                )))?;

            if let Ok(message) = outlook_pst::messaging::message::UnicodeMessage::read(
                store.clone(),
                &message_entry_id,
                Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x0E06, 0x1000, 0x1013, 0x1009]),
            ) {
                let props = message.properties();

                let get_str = |id: u16| -> String {
                    props
                        .get(id)
                        .and_then(|v| match v {
                            PropertyValue::String8(s) => Some(s.to_string()),
                            PropertyValue::Unicode(s) => Some(s.to_string()),
                            _ => None,
                        })
                        .unwrap_or_default()
                };

                let from = get_str(0x0C1A);
                let to = get_str(0x0E04);
                let cc = get_str(0x0E02);
                let subject = get_str(0x0037);

                let get_str_prop = |id: u16| -> Option<String> {
                    props.get(id).and_then(|v| match v {
                        PropertyValue::String8(s) => Some(s.to_string()),
                        PropertyValue::Unicode(s) => Some(s.to_string()),
                        _ => None,
                    })
                };
                let body = if let Some(s) = get_str_prop(0x1000) {
                    s
                } else if let Some(html) = get_str_prop(0x1013) {
                    html_to_text(&html)
                } else if let Some(PropertyValue::Binary(rtf)) = props.get(0x1009) {
                    rtf_compressed_to_text(rtf.buffer()).unwrap_or_default()
                } else {
                    String::new()
                };

                let matches = [&from, &to, &cc, &body]
                    .iter()
                    .any(|s| s.to_ascii_lowercase().contains(query_lower));

                if matches {
                    let date = props
                        .get(0x0E06)
                        .and_then(|v| match v {
                            PropertyValue::Time(t) => Some(filetime_to_string(*t)),
                            _ => None,
                        })
                        .unwrap_or_default();

                    records.push(SearchRecord {
                        folder: folder_name.clone(),
                        subject,
                        from,
                        to,
                        cc,
                        date,
                    });
                }
            }
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for subfolder_row in hierarchy_table.rows_matrix() {
            let subfolder_entry_id = store
                .properties()
                .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                    subfolder_row.id(),
                )))?;
            let subfolder = outlook_pst::messaging::folder::UnicodeFolder::read(
                store.clone(),
                &subfolder_entry_id,
            )?;
            collect_search_matches(store.clone(), &subfolder, query_lower, records)?;
        }
    }

    Ok(())
}

fn search_traverse_folders(
    store: Rc<UnicodeStore>,
    folder: &outlook_pst::messaging::folder::UnicodeFolder,
    query_lower: &str,
) -> Result<usize, Box<dyn std::error::Error>> {
    let mut match_count = 0;

    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    if let Some(contents_table) = folder.contents_table() {
        for message_row in contents_table.rows_matrix() {
            let message_entry_id = store
                .properties()
                .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                    message_row.id(),
                )))?;

            if let Ok(message) = outlook_pst::messaging::message::UnicodeMessage::read(
                store.clone(),
                &message_entry_id,
                Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x0E06, 0x1000, 0x1013, 0x1009]),
            ) {
                let props = message.properties();

                let get_str = |id: u16| -> String {
                    props
                        .get(id)
                        .and_then(|v| match v {
                            PropertyValue::String8(s) => Some(s.to_string()),
                            PropertyValue::Unicode(s) => Some(s.to_string()),
                            _ => None,
                        })
                        .unwrap_or_default()
                };

                let from = get_str(0x0C1A);
                let to = get_str(0x0E04);
                let cc = get_str(0x0E02);
                let subject = get_str(0x0037);

                let get_str_prop = |id: u16| -> Option<String> {
                    props.get(id).and_then(|v| match v {
                        PropertyValue::String8(s) => Some(s.to_string()),
                        PropertyValue::Unicode(s) => Some(s.to_string()),
                        _ => None,
                    })
                };
                let body = if let Some(s) = get_str_prop(0x1000) {
                    s
                } else if let Some(html) = get_str_prop(0x1013) {
                    html_to_text(&html)
                } else if let Some(PropertyValue::Binary(rtf)) = props.get(0x1009) {
                    rtf_compressed_to_text(rtf.buffer()).unwrap_or_default()
                } else {
                    String::new()
                };

                let matches = [&from, &to, &cc, &body]
                    .iter()
                    .any(|s| s.to_ascii_lowercase().contains(query_lower));

                if matches {
                    let date = props
                        .get(0x0E06)
                        .and_then(|v| match v {
                            PropertyValue::Time(t) => Some(filetime_to_string(*t)),
                            _ => None,
                        })
                        .unwrap_or_default();

                    println!("Folder:  {}", folder_name);
                    println!("Subject: {}", subject);
                    println!("From:    {}", from);
                    println!("To:      {}", to);
                    if !cc.is_empty() {
                        println!("CC:      {}", cc);
                    }
                    println!("Date:    {}", date);
                    println!("---");
                    match_count += 1;
                }
            }
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for subfolder_row in hierarchy_table.rows_matrix() {
            let subfolder_entry_id = store
                .properties()
                .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                    subfolder_row.id(),
                )))?;
            let subfolder = outlook_pst::messaging::folder::UnicodeFolder::read(
                store.clone(),
                &subfolder_entry_id,
            )?;
            match_count += search_traverse_folders(store.clone(), &subfolder, query_lower)?;
        }
    }

    Ok(match_count)
}

fn run_search(browser: &PstBrowser, query: &str) -> Vec<SearchResultItem> {
    let query_lower = query.to_ascii_lowercase();
    let mut results = Vec::new();
    collect_search_results(
        Rc::clone(&browser.store),
        &browser.root_folder,
        &query_lower,
        &mut results,
    );
    results
}

fn collect_search_results(
    store: Rc<UnicodeStore>,
    folder: &UnicodeFolder,
    query_lower: &str,
    results: &mut Vec<SearchResultItem>,
) {
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    if let Some(contents_table) = folder.contents_table() {
        for message_row in contents_table.rows_matrix() {
            let row_id = u32::from(message_row.id());
            let entry_id = match store
                .properties()
                .make_entry_id(NodeId::from(row_id))
            {
                Ok(id) => id,
                Err(_) => continue,
            };

            let message = match UnicodeMessage::read(
                Rc::clone(&store),
                &entry_id,
                Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x0039, 0x0E06, 0x1000, 0x1013, 0x1009]),
            ) {
                Ok(m) => m,
                Err(_) => continue,
            };

            let props = message.properties();
            let get_str = |id: u16| -> String {
                props
                    .get(id)
                    .and_then(|v| match v {
                        PropertyValue::String8(s) => Some(s.to_string()),
                        PropertyValue::Unicode(s) => Some(s.to_string()),
                        _ => None,
                    })
                    .unwrap_or_default()
            };

            let from = get_str(0x0C1A);
            let to = get_str(0x0E04);
            let cc = get_str(0x0E02);
            let subject = get_str(0x0037);

            let get_str_prop = |id: u16| -> Option<String> {
                props.get(id).and_then(|v| match v {
                    PropertyValue::String8(s) => Some(s.to_string()),
                    PropertyValue::Unicode(s) => Some(s.to_string()),
                    _ => None,
                })
            };
            let body = if let Some(s) = get_str_prop(0x1000) {
                s
            } else if let Some(html) = get_str_prop(0x1013) {
                html_to_text(&html)
            } else if let Some(PropertyValue::Binary(rtf)) = props.get(0x1009) {
                rtf_compressed_to_text(rtf.buffer()).unwrap_or_default()
            } else {
                String::new()
            };

            let matches = [&from, &to, &cc, &body]
                .iter()
                .any(|s| s.to_ascii_lowercase().contains(query_lower));

            if matches {
                let date = props
                    .get(0x0039)
                    .or_else(|| props.get(0x0E06))
                    .and_then(|v| match v {
                        PropertyValue::Time(t) => Some(filetime_to_string(*t)),
                        _ => None,
                    })
                    .unwrap_or_default();
                results.push(SearchResultItem {
                    folder_name: folder_name.clone(),
                    row_data: MessageRow { from, to, cc, subject, date },
                    row_id,
                });
            }
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for subfolder_row in hierarchy_table.rows_matrix() {
            let Ok(entry_id) = store
                .properties()
                .make_entry_id(NodeId::from(u32::from(subfolder_row.id())))
            else {
                continue;
            };
            let Ok(subfolder) = UnicodeFolder::read(Rc::clone(&store), &entry_id) else {
                continue;
            };
            collect_search_results(Rc::clone(&store), &subfolder, query_lower, results);
        }
    }
}

// ── SQLite Export ────────────────────────────────────────────────────────

fn create_export_schema(conn: &Connection) -> Result<(), rusqlite::Error> {
    conn.execute_batch(
        "
        CREATE TABLE folders (
            id        INTEGER PRIMARY KEY,
            parent_id INTEGER REFERENCES folders(id),
            name      TEXT NOT NULL,
            path      TEXT NOT NULL
        );

        CREATE TABLE messages (
            id               INTEGER PRIMARY KEY,
            folder_id        INTEGER NOT NULL REFERENCES folders(id),
            message_class    TEXT NOT NULL,
            subject          TEXT,
            sender           TEXT,
            to_recipients    TEXT,
            cc_recipients    TEXT,
            submit_time      TEXT,
            delivery_time    TEXT,
            body_text        TEXT,
            body_html        TEXT,
            body_rtf         BLOB,
            attachment_count INTEGER DEFAULT 0
        );

        CREATE TABLE attachments (
            id           INTEGER PRIMARY KEY,
            message_id   INTEGER NOT NULL REFERENCES messages(id),
            filename     TEXT,
            content_type TEXT,
            size         INTEGER,
            data         BLOB
        );

        CREATE INDEX idx_messages_folder ON messages(folder_id);
        CREATE INDEX idx_messages_class  ON messages(message_class);
        CREATE INDEX idx_messages_sender ON messages(sender);
        CREATE INDEX idx_messages_submit ON messages(submit_time);
        CREATE INDEX idx_attachments_msg ON attachments(message_id);
        ",
    )
}

fn filetime_to_iso(ticks: i64) -> Option<String> {
    const EPOCH_DIFF_SECS: i64 = 11_644_473_600;
    let secs = ticks / 10_000_000 - EPOCH_DIFF_SECS;
    let nanos = ((ticks % 10_000_000) * 100) as u32;
    match Utc.timestamp_opt(secs, nanos) {
        chrono::LocalResult::Single(dt) => Some(dt.format("%Y-%m-%dT%H:%M:%SZ").to_string()),
        _ => None,
    }
}

fn export_folder(
    store: Rc<UnicodeStore>,
    folder: &UnicodeFolder,
    parent_folder_id: Option<i64>,
    path_prefix: &str,
    conn: &Connection,
    counts: &mut (usize, usize),
) -> Result<(), Box<dyn std::error::Error>> {
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());
    let path = if path_prefix.is_empty() {
        folder_name.clone()
    } else {
        format!("{}/{}", path_prefix, folder_name)
    };

    conn.execute(
        "INSERT INTO folders (parent_id, name, path) VALUES (?1, ?2, ?3)",
        params![parent_folder_id, &folder_name, &path],
    )?;
    let folder_id = conn.last_insert_rowid();
    counts.0 += 1;

    if let Some(contents_table) = folder.contents_table() {
        for row in contents_table.rows_matrix() {
            let row_id = u32::from(row.id());
            let entry_id = match store.properties().make_entry_id(NodeId::from(row_id)) {
                Ok(id) => id,
                Err(_) => continue,
            };

            let message = match UnicodeMessage::read(
                Rc::clone(&store),
                &entry_id,
                Some(&[
                    0x0037, 0x001A, 0x0039, 0x0C1A, 0x0E02, 0x0E04, 0x0E06, 0x0E13, 0x1000,
                    0x1009, 0x1013,
                ]),
            ) {
                Ok(m) => m,
                Err(_) => continue,
            };

            let props = message.properties();

            let get_str = |id: u16| -> Option<String> {
                props.get(id).and_then(|v| match v {
                    PropertyValue::String8(s) => Some(s.to_string()),
                    PropertyValue::Unicode(s) => Some(s.to_string()),
                    _ => None,
                })
            };

            let message_class = get_str(0x001A)
                .map(|s| s.to_ascii_uppercase())
                .unwrap_or_else(|| "IPM.NOTE".to_string());
            let subject = get_str(0x0037);
            let sender = get_str(0x0C1A);
            let to_recipients = get_str(0x0E04);
            let cc_recipients = get_str(0x0E02);

            let submit_time = props.get(0x0039).and_then(|v| match v {
                PropertyValue::Time(t) => filetime_to_iso(*t),
                _ => None,
            });
            let delivery_time = props.get(0x0E06).and_then(|v| match v {
                PropertyValue::Time(t) => filetime_to_iso(*t),
                _ => None,
            });

            let body_text = get_str(0x1000);

            let body_html: Option<String> = props.get(0x1013).and_then(|v| match v {
                PropertyValue::Binary(b) => {
                    let s = String::from_utf8_lossy(b.buffer());
                    if s.trim_start().starts_with('<') {
                        Some(s.into_owned())
                    } else {
                        None
                    }
                }
                PropertyValue::String8(s) => Some(s.to_string()),
                PropertyValue::Unicode(s) => Some(s.to_string()),
                _ => None,
            });

            let body_rtf: Option<Vec<u8>> = props.get(0x1009).and_then(|v| match v {
                PropertyValue::Binary(b) => Some(b.buffer().to_vec()),
                _ => None,
            });

            let attachment_count: i32 = props
                .get(0x0E13)
                .and_then(|v| match v {
                    PropertyValue::Integer32(n) => Some(*n),
                    _ => None,
                })
                .unwrap_or(0);

            conn.execute(
                "INSERT INTO messages (folder_id, message_class, subject, sender,
                    to_recipients, cc_recipients, submit_time, delivery_time,
                    body_text, body_html, body_rtf, attachment_count)
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12)",
                params![
                    folder_id,
                    &message_class,
                    &subject,
                    &sender,
                    &to_recipients,
                    &cc_recipients,
                    &submit_time,
                    &delivery_time,
                    &body_text,
                    &body_html,
                    &body_rtf,
                    attachment_count,
                ],
            )?;
            counts.1 += 1;
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for row in hierarchy_table.rows_matrix() {
            let Ok(entry_id) = store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
            else {
                continue;
            };
            let Ok(subfolder) = UnicodeFolder::read(Rc::clone(&store), &entry_id) else {
                continue;
            };
            export_folder(
                Rc::clone(&store),
                &subfolder,
                Some(folder_id),
                &path,
                conn,
                counts,
            )?;
        }
    }

    Ok(())
}

fn export_pst(
    file_path: &PathBuf,
    output: Option<&PathBuf>,
) -> Result<(), Box<dyn std::error::Error>> {
    let db_path = match output {
        Some(p) => p.clone(),
        None => {
            let stem = file_path
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("export");
            PathBuf::from(format!("{}.db", stem))
        }
    };

    if db_path.exists() {
        return Err(format!("Output file already exists: {:?}", db_path).into());
    }

    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = UnicodeFolder::read(Rc::clone(&store), &ipm_sub_tree_entry_id)?;

    let conn = Connection::open(&db_path)?;
    conn.execute_batch("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;")?;
    create_export_schema(&conn)?;

    let mut counts: (usize, usize) = (0, 0);
    conn.execute_batch("BEGIN TRANSACTION;")?;
    export_folder(
        Rc::clone(&store),
        &ipm_subtree_folder,
        None,
        "",
        &conn,
        &mut counts,
    )?;
    conn.execute_batch("COMMIT;")?;

    println!("Exported to {:?}", db_path);
    println!("  Folders:  {}", counts.0);
    println!("  Messages: {}", counts.1);
    Ok(())
}

fn collect_all_messages(
    store: Rc<UnicodeStore>,
    folder: &UnicodeFolder,
    results: &mut Vec<(String, u32)>,
) {
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    if let Some(contents_table) = folder.contents_table() {
        for row in contents_table.rows_matrix() {
            results.push((folder_name.clone(), u32::from(row.id())));
        }
    }

    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for row in hierarchy_table.rows_matrix() {
            let Ok(entry_id) = store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
            else {
                continue;
            };
            let Ok(subfolder) = UnicodeFolder::read(Rc::clone(&store), &entry_id) else {
                continue;
            };
            collect_all_messages(Rc::clone(&store), &subfolder, results);
        }
    }
}

fn traverse_folder_hierarchy(
    store: Rc<UnicodeStore>,
    folder: &outlook_pst::messaging::folder::UnicodeFolder,
    limit: usize,
    printed: &mut usize,
) -> Result<usize, Box<dyn std::error::Error>> {
    let mut email_count = 0;

    // Get the folder name
    let folder_name = folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Unknown".to_string());

    println!("\nFolder: {}", folder_name);

    // Process messages in this folder
    if let Some(contents_table) = folder.contents_table() {
        let messages: Vec<_> = contents_table.rows_matrix().collect();
        let message_count = messages.len();

        println!("  Contains {} messages", message_count);

        // Process each message
        for message_row in messages {
            // Stop printing messages if the limit has been reached
            if limit > 0 && *printed >= limit {
                break;
            }

            let message_entry_id =
                store
                    .properties()
                    .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                        message_row.id(),
                    )))?;

            if let Ok(message) = outlook_pst::messaging::message::UnicodeMessage::read(
                store.clone(),
                &message_entry_id,
                Some(&[0x0037, 0x0C1A, 0x0E06]), // Subject, Sender, Received Time
            ) {
                // Extract message details
                let properties = message.properties();

                let subject = properties
                    .get(0x0037)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::String8(s) => {
                            Some(s.to_string())
                        }
                        outlook_pst::ltp::prop_context::PropertyValue::Unicode(s) => {
                            Some(s.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "No Subject".to_string());

                let sender = properties
                    .get(0x0C1A)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::String8(s) => {
                            Some(s.to_string())
                        }
                        outlook_pst::ltp::prop_context::PropertyValue::Unicode(s) => {
                            Some(s.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "Unknown Sender".to_string());

                let received_time = properties
                    .get(0x0E06)
                    .and_then(|v| match v {
                        outlook_pst::ltp::prop_context::PropertyValue::Time(t) => {
                            Some(t.to_string())
                        }
                        _ => None,
                    })
                    .unwrap_or_else(|| "Unknown Date".to_string());

                println!("  - Subject: {}", subject);
                println!("    From: {}", sender);
                println!("    Date: {}", received_time);
                println!("    ---");

                email_count += 1;
                *printed += 1;
            }
        }
    }

    // Recursively traverse subfolders (stop if limit reached)
    if let Some(hierarchy_table) = folder.hierarchy_table() {
        for subfolder_row in hierarchy_table.rows_matrix() {
            if limit > 0 && *printed >= limit {
                break;
            }
            let subfolder_entry_id =
                store
                    .properties()
                    .make_entry_id(outlook_pst::ndb::node_id::NodeId::from(u32::from(
                        subfolder_row.id(),
                    )))?;
            let subfolder = outlook_pst::messaging::folder::UnicodeFolder::read(
                store.clone(),
                &subfolder_entry_id,
            )?;
            email_count += traverse_folder_hierarchy(store.clone(), &subfolder, limit, printed)?;
        }
    }

    Ok(email_count)
}

// TUI Data Structures
struct PstBrowser {
    store: Rc<UnicodeStore>,
    root_folder: Rc<UnicodeFolder>,
}

#[derive(Clone, Default)]
struct MessageRow {
    from: String,
    to: String,
    cc: String,
    subject: String,
    date: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ColumnId {
    From,
    To,
    Cc,
    Subject,
    Date,
}

struct ColumnConfig {
    id: ColumnId,
    label: &'static str,
    width: Constraint,
    visible: bool,
}

fn default_columns() -> Vec<ColumnConfig> {
    vec![
        ColumnConfig { id: ColumnId::From,    label: "From",    width: Constraint::Percentage(20), visible: true },
        ColumnConfig { id: ColumnId::To,      label: "To",      width: Constraint::Percentage(20), visible: true },
        ColumnConfig { id: ColumnId::Cc,      label: "CC",      width: Constraint::Percentage(15), visible: false },
        ColumnConfig { id: ColumnId::Subject, label: "Subject", width: Constraint::Percentage(40), visible: true },
        ColumnConfig { id: ColumnId::Date,    label: "Date",    width: Constraint::Percentage(20), visible: true },
    ]
}

struct SearchResultItem {
    folder_name: String,
    row_data: MessageRow,
    row_id: u32,
}

#[derive(Default)]
struct MessageHeaders {
    from: String,
    to: String,
    cc: String,
    subject: String,
    date: String,
}

#[derive(PartialEq)]
enum ActivePane {
    Messages,
    Preview,
}

struct AppState {
    exit: bool,
    /// All message row IDs across every folder, collected at startup.
    all_row_ids: Vec<u32>,
    /// Folder name for each entry in all_row_ids.
    all_folder_names: Vec<String>,
    /// Row IDs for the current view (all messages, or search results).
    message_row_ids: Vec<u32>,
    /// Folder name for each entry in message_row_ids.
    message_folder_names: Vec<String>,
    /// Lazily loaded row data; None = not yet fetched.
    message_rows: Vec<Option<MessageRow>>,
    message_table_state: TableState,
    /// Column configuration for the message table.
    columns: Vec<ColumnConfig>,
    /// Height of the message list area as of the last draw — used to size the load window.
    message_list_height: usize,
    current_message_content: String,
    current_headers: MessageHeaders,
    active_pane: ActivePane,
    preview_scroll: u16,
    /// Debug event log; None if debug mode not enabled.
    debug_log: Option<Vec<String>>,
    /// Transient status bar message (cleared on next keypress).
    status_message: Option<String>,
    /// Current text in the search bar.
    search_input: String,
    /// Whether keyboard input is being captured by the search bar.
    search_bar_active: bool,
    /// Whether we are currently showing search results instead of all messages.
    search_mode: bool,
    /// Set when Enter is pressed in search bar; cleared after search completes.
    search_pending: bool,
    /// When the current/last search started (for elapsed-time display).
    search_start: Option<Instant>,
}

impl PstBrowser {
    fn new(store: Rc<UnicodeStore>, root_folder: Rc<UnicodeFolder>) -> Self {
        Self { store, root_folder }
    }
}

impl AppState {
    fn new(browser: &PstBrowser, debug: bool) -> Self {
        let mut all_messages: Vec<(String, u32)> = Vec::new();
        collect_all_messages(
            Rc::clone(&browser.store),
            &browser.root_folder,
            &mut all_messages,
        );
        let all_row_ids: Vec<u32> = all_messages.iter().map(|(_, id)| *id).collect();
        let all_folder_names: Vec<String> = all_messages.iter().map(|(n, _)| n.clone()).collect();
        let n = all_row_ids.len();

        let mut message_table_state = TableState::default();
        if n > 0 {
            message_table_state.select(Some(0));
        }

        Self {
            exit: false,
            message_row_ids: all_row_ids.clone(),
            message_folder_names: all_folder_names.clone(),
            all_row_ids,
            all_folder_names,
            message_rows: vec![None; n],
            message_table_state,
            columns: default_columns(),
            message_list_height: 20,
            current_message_content: if n == 0 {
                "No messages found".to_string()
            } else {
                "Select a message to view its content".to_string()
            },
            current_headers: MessageHeaders::default(),
            active_pane: ActivePane::Messages,
            preview_scroll: 0,
            debug_log: if debug { Some(vec![]) } else { None },
            status_message: None,
            search_input: String::new(),
            search_bar_active: false,
            search_mode: false,
            search_pending: false,
            search_start: None,
        }
    }

    /// Load row data for the visible window around the current scroll offset.
    fn load_visible_rows(&mut self, browser: &PstBrowser) {
        let offset = self.message_table_state.offset();
        let end = (offset + self.message_list_height + 5).min(self.message_row_ids.len());
        for i in offset..end {
            if self.message_rows[i].is_none() {
                let row = browser
                    .store
                    .properties()
                    .make_entry_id(NodeId::from(self.message_row_ids[i]))
                    .ok()
                    .and_then(|eid| {
                        UnicodeMessage::read(
                            Rc::clone(&browser.store),
                            &eid,
                            Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x0039, 0x0E06]),
                        )
                        .ok()
                    })
                    .map(|msg| {
                        let props = msg.properties();
                        let get_str = |id: u16| -> String {
                            props
                                .get(id)
                                .and_then(|v| match v {
                                    PropertyValue::String8(s) => Some(s.to_string()),
                                    PropertyValue::Unicode(s) => Some(s.to_string()),
                                    _ => None,
                                })
                                .unwrap_or_default()
                        };
                        let date = props
                            .get(0x0039)
                            .or_else(|| props.get(0x0E06))
                            .and_then(|v| match v {
                                PropertyValue::Time(t) => Some(filetime_to_string(*t)),
                                _ => None,
                            })
                            .unwrap_or_default();
                        MessageRow {
                            from: get_str(0x0C1A),
                            to: get_str(0x0E04),
                            cc: get_str(0x0E02),
                            subject: get_str(0x0037),
                            date,
                        }
                    })
                    .unwrap_or_else(|| MessageRow {
                        subject: "(no subject)".to_string(),
                        ..Default::default()
                    });
                self.message_rows[i] = Some(row);
            }
        }
    }

    fn log_event(&mut self, msg: &str) {
        if let Some(log) = &mut self.debug_log {
            log.push(msg.to_string());
        }
    }

    fn restore_all_messages(&mut self) {
        let n = self.all_row_ids.len();
        self.message_row_ids = self.all_row_ids.clone();
        self.message_folder_names = self.all_folder_names.clone();
        self.message_rows = vec![None; n];
        self.search_mode = false;
        self.search_input.clear();
        self.message_table_state = TableState::default();
        if n > 0 { self.message_table_state.select(Some(0)); }
        self.current_headers = MessageHeaders::default();
        self.current_message_content = "Select a message to view its content".to_string();
        self.preview_scroll = 0;
    }

    fn select_message(&mut self, browser: &PstBrowser, index: usize) {
        if let Some(&row_id) = self.message_row_ids.get(index) {
            let entry_id = browser
                .store
                .properties()
                .make_entry_id(NodeId::from(row_id))
                .ok();

            let message_result = entry_id.and_then(|eid| {
                UnicodeMessage::read(
                    Rc::clone(&browser.store),
                    &eid,
                    Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x0039, 0x0E06, 0x1000, 0x1013, 0x1009]),
                ).ok()
            });
            if message_result.is_none() {
                self.current_message_content =
                    "(This item type cannot be displayed — not a standard email message)".to_string();
                self.current_headers = MessageHeaders::default();
                self.preview_scroll = 0;
            }
            if let Some(message) = message_result
            {
                let props = message.properties();

                let get_str = |id: u16| -> String {
                    props
                        .get(id)
                        .and_then(|v| match v {
                            PropertyValue::String8(s) => Some(s.to_string()),
                            PropertyValue::Unicode(s) => Some(s.to_string()),
                            _ => None,
                        })
                        .unwrap_or_default()
                };

                let date = props
                    .get(0x0039)
                    .or_else(|| props.get(0x0E06))
                    .and_then(|v| match v {
                        PropertyValue::Time(t) => Some(filetime_to_string(*t)),
                        _ => None,
                    })
                    .unwrap_or_default();

                self.current_headers = MessageHeaders {
                    subject: get_str(0x0037),
                    from: get_str(0x0C1A),
                    to: get_str(0x0E04),
                    cc: get_str(0x0E02),
                    date,
                };

                self.current_message_content = props
                    .get(0x1000)
                    .and_then(|value| match value {
                        PropertyValue::String8(s) => Some(s.to_string()),
                        PropertyValue::Unicode(s) => Some(s.to_string()),
                        _ => None,
                    })
                    .or_else(|| {
                        props.get(0x1013).and_then(|value| match value {
                            PropertyValue::Binary(b) => {
                                // Sanity-check: real HTML starts with '<' (possibly after BOM/whitespace).
                                // If it doesn't, it's likely compressed/binary — skip it.
                                let s = String::from_utf8_lossy(b.buffer());
                                if s.trim_start().starts_with('<') {
                                    Some(html_to_text(&s))
                                } else {
                                    None
                                }
                            }
                            PropertyValue::String8(s) => Some(html_to_text(&s.to_string())),
                            PropertyValue::Unicode(s) => Some(html_to_text(&s.to_string())),
                            _ => None,
                        })
                    })
                    .or_else(|| {
                        props.get(0x1009).and_then(|value| match value {
                            PropertyValue::Binary(b) => rtf_compressed_to_text(b.buffer()),
                            _ => None,
                        })
                    })
                    .unwrap_or_else(|| "No message content available".to_string());
                self.preview_scroll = 0;
            }
        }
    }
}

fn browse_pst(file_path: &PathBuf, debug: bool) -> Result<(), Box<dyn std::error::Error>> {
    // Open the PST file
    let pst = UnicodePstFile::open(file_path)?;
    let pst_rc = Rc::new(pst);
    let store = Rc::new(UnicodeStore::read(Rc::clone(&pst_rc))?);

    // Get the IPM subtree (where emails are stored)
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let root_folder = UnicodeFolder::read(Rc::clone(&store), &ipm_sub_tree_entry_id)?;

    let browser = PstBrowser::new(Rc::clone(&store), root_folder);

    // Setup terminal
    match enable_raw_mode() {
        Ok(_) => {
            let mut stdout = io::stdout();
            if execute!(stdout, EnterAlternateScreen).is_ok() {
                let backend = CrosstermBackend::new(stdout);
                if let Ok(mut terminal) = Terminal::new(backend) {
                    let mut app_state = AppState::new(&browser, debug);

                    // Main loop
                    while !app_state.exit {
                        // Load row data for the currently visible window before drawing.
                        app_state.load_visible_rows(&browser);

                        if let Err(e) =
                            terminal.draw(|frame| draw_ui(frame, &browser, &mut app_state))
                        {
                            eprintln!("Error drawing UI: {}", e);
                            break;
                        }

                        // Run any pending search after drawing so the "Searching..." status
                        // is visible for at least one frame before blocking.
                        if app_state.search_pending {
                            app_state.search_pending = false;
                            let results = run_search(&browser, &app_state.search_input);
                            let elapsed = app_state.search_start
                                .take()
                                .map(|t| t.elapsed().as_secs())
                                .unwrap_or(0);
                            let n = results.len();
                            app_state.search_mode = true;
                            app_state.message_row_ids = results.iter().map(|r| r.row_id).collect();
                            app_state.message_folder_names = results.iter().map(|r| r.folder_name.clone()).collect();
                            app_state.message_rows = results.iter().map(|r| Some(r.row_data.clone())).collect();
                            app_state.message_table_state = TableState::default();
                            if n > 0 {
                                app_state.message_table_state.select(Some(0));
                                app_state.select_message(&browser, 0);
                            } else {
                                app_state.current_headers = MessageHeaders::default();
                                app_state.current_message_content = "No messages match the search query".to_string();
                            }
                            app_state.preview_scroll = 0;
                            app_state.active_pane = ActivePane::Messages;
                            app_state.status_message = Some(format!(
                                "Found {} result{} ({}s)",
                                n, if n == 1 { "" } else { "s" }, elapsed
                            ));
                            continue; // redraw immediately to show results
                        }

                        if let Err(e) = handle_events(&mut app_state, &browser) {
                            eprintln!("Error handling events: {}", e);
                            break;
                        }
                    }

                    // Write debug log if enabled
                    if let Some(log) = &app_state.debug_log {
                        let content = log.join("\n") + "\n";
                        let _ = std::fs::write("pstexplorer-debug.log", content);
                    }

                    // Cleanup
                    let _ = disable_raw_mode();
                    let _ = execute!(terminal.backend_mut(), LeaveAlternateScreen);
                }
            }
        }
        Err(e) => {
            eprintln!("Could not enable raw terminal mode: {}", e);
            eprintln!(
                "This typically happens when running in an environment that doesn't support terminal UI (like some IDEs or non-interactive shells)."
            );
            eprintln!("Please run this command in a proper terminal emulator.");
            eprintln!();
            eprintln!("For now, here's the basic information about the PST file:");

            // Fall back to showing basic info
            list_emails(file_path, None, 0)?;
        }
    }

    Ok(())
}

fn draw_ui(frame: &mut ratatui::Frame, _browser: &PstBrowser, state: &mut AppState) {
    let layout = Layout::default()
        .direction(ratatui::layout::Direction::Vertical)
        .constraints([
            Constraint::Length(1), // Search bar
            Constraint::Min(0),    // Main content
            Constraint::Length(1), // Status bar
        ])
        .split(frame.area());

    draw_search_bar(frame, state, layout[0]);

    let main_layout = Layout::default()
        .direction(ratatui::layout::Direction::Vertical)
        .constraints([
            Constraint::Percentage(35), // Message list
            Constraint::Percentage(65), // Message preview
        ])
        .split(layout[1]);

    draw_message_list(frame, state, main_layout[0]);
    draw_message_preview(frame, state, main_layout[1]);

    let status_text = if let Some(msg) = &state.status_message {
        msg.clone()
    } else if state.search_bar_active {
        " [Search] type to search  Enter: run  Esc: cancel".to_string()
    } else {
        match state.active_pane {
            ActivePane::Messages => " [Messages] j/k: navigate  Enter/Tab: preview  /: search  Esc: clear search  q: quit".to_string(),
            ActivePane::Preview => " [Preview] j/k: scroll  Tab: → messages  /: search  Esc: clear search  q: quit".to_string(),
        }
    };
    let status_style = if state.status_message.is_some() {
        Style::default().fg(ratatui::style::Color::Green)
    } else {
        Style::default().fg(ratatui::style::Color::DarkGray)
    };
    let status = ratatui::widgets::Paragraph::new(status_text).style(status_style);
    frame.render_widget(status, layout[2]);
}

fn draw_search_bar(frame: &mut ratatui::Frame, state: &AppState, area: Rect) {
    let (label_style, input_style, cursor_style) = if state.search_bar_active {
        (
            Style::default().fg(ratatui::style::Color::Cyan).add_modifier(Modifier::BOLD),
            Style::default().fg(ratatui::style::Color::White),
            Style::default().fg(ratatui::style::Color::Cyan).add_modifier(Modifier::BOLD),
        )
    } else if state.search_mode {
        (
            Style::default().fg(ratatui::style::Color::Yellow).add_modifier(Modifier::BOLD),
            Style::default().fg(ratatui::style::Color::White),
            Style::default(),
        )
    } else {
        (
            Style::default().fg(ratatui::style::Color::DarkGray),
            Style::default().fg(ratatui::style::Color::DarkGray),
            Style::default(),
        )
    };

    let mut spans = vec![
        Span::styled(" Search: ", label_style),
        Span::styled(state.search_input.clone(), input_style),
    ];
    if state.search_bar_active {
        spans.push(Span::styled("▋", cursor_style));
    } else if state.search_mode {
        let count = state.message_row_ids.len();
        spans.push(Span::styled(
            format!("  ({} result{})", count, if count == 1 { "" } else { "s" }),
            Style::default().fg(ratatui::style::Color::Yellow),
        ));
    } else if state.search_input.is_empty() {
        spans.push(Span::styled("Press / to search", Style::default().fg(ratatui::style::Color::DarkGray)));
    }

    frame.render_widget(
        Paragraph::new(Line::from(spans)),
        area,
    );
}

fn draw_message_list(frame: &mut ratatui::Frame, state: &mut AppState, area: Rect) {
    // Subtract 3 for border top + header row + border bottom.
    state.message_list_height = area.height.saturating_sub(3) as usize;

    let count = state.message_row_ids.len();
    let selected_num = state.message_table_state.selected().map(|i| i + 1).unwrap_or(0);
    let title = if state.search_mode {
        format!("Search Results ({}/{})", selected_num, count)
    } else {
        format!("Messages ({}/{})", selected_num, count)
    };

    let visible_cols: Vec<&ColumnConfig> = state.columns.iter().filter(|c| c.visible).collect();

    let header_cells: Vec<Cell> = visible_cols
        .iter()
        .map(|col| {
            Cell::from(col.label).style(
                Style::default()
                    .fg(ratatui::style::Color::Yellow)
                    .add_modifier(Modifier::BOLD),
            )
        })
        .collect();
    let header = Row::new(header_cells).bottom_margin(0);

    let empty = MessageRow::default();
    let rows: Vec<Row> = (0..count)
        .map(|i| {
            let row_data = state.message_rows[i].as_ref().unwrap_or(&empty);
            let cells: Vec<Cell> = visible_cols
                .iter()
                .map(|col| {
                    let val = match col.id {
                        ColumnId::From => &row_data.from,
                        ColumnId::To => &row_data.to,
                        ColumnId::Cc => &row_data.cc,
                        ColumnId::Subject => &row_data.subject,
                        ColumnId::Date => &row_data.date,
                    };
                    Cell::from(if val.is_empty() && matches!(col.id, ColumnId::Subject) {
                        "\u{2026}" // ellipsis for loading
                    } else {
                        val.as_str()
                    })
                })
                .collect();
            Row::new(cells)
        })
        .collect();

    let widths: Vec<Constraint> = visible_cols.iter().map(|c| c.width).collect();

    let border_style = if state.active_pane == ActivePane::Messages {
        Style::default().fg(ratatui::style::Color::Cyan)
    } else {
        Style::default()
    };

    let table = Table::new(rows, &widths)
        .header(header)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(border_style)
                .title(title),
        )
        .row_highlight_style(
            Style::default()
                .fg(ratatui::style::Color::Green)
                .add_modifier(Modifier::BOLD),
        );

    frame.render_stateful_widget(table, area, &mut state.message_table_state);
}

fn draw_message_preview(frame: &mut ratatui::Frame, state: &AppState, area: Rect) {
    let border_style = if state.active_pane == ActivePane::Preview {
        Style::default().fg(ratatui::style::Color::Cyan)
    } else {
        Style::default()
    };

    let label_style = Style::default()
        .fg(ratatui::style::Color::Yellow)
        .add_modifier(Modifier::BOLD);
    let value_style = Style::default();

    let h = &state.current_headers;
    let mut lines: Vec<Line> = vec![
        Line::from(vec![
            Span::styled("From:    ", label_style),
            Span::styled(h.from.clone(), value_style),
        ]),
        Line::from(vec![
            Span::styled("To:      ", label_style),
            Span::styled(h.to.clone(), value_style),
        ]),
        Line::from(vec![
            Span::styled("CC:      ", label_style),
            Span::styled(h.cc.clone(), value_style),
        ]),
        Line::from(vec![
            Span::styled("Subject: ", label_style),
            Span::styled(h.subject.clone(), value_style),
        ]),
        Line::from(vec![
            Span::styled("Date:    ", label_style),
            Span::styled(h.date.clone(), value_style),
        ]),
        Line::from("─".repeat(area.width.saturating_sub(2) as usize)),
    ];

    for line in state.current_message_content.lines() {
        lines.push(Line::from(line.to_string()));
    }

    let preview = Paragraph::new(Text::from(lines))
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(border_style)
                .title("Message Preview"),
        )
        .wrap(ratatui::widgets::Wrap { trim: false })
        .scroll((state.preview_scroll, 0));

    frame.render_widget(preview, area);
}

fn handle_events(
    state: &mut AppState,
    browser: &PstBrowser,
) -> Result<(), Box<dyn std::error::Error>> {
    if event::poll(std::time::Duration::from_millis(100))?
        && let Event::Key(key) = event::read()?
        && key.kind == KeyEventKind::Press
    {
        state.status_message = None;

        let pane_name = match state.active_pane {
            ActivePane::Messages => "Messages",
            ActivePane::Preview => "Preview",
        };
        let key_str = match key.code {
            KeyCode::Char(c) => format!("'{}'", c),
            KeyCode::Enter => "Enter".to_string(),
            KeyCode::Tab => "Tab".to_string(),
            KeyCode::Esc => "Esc".to_string(),
            KeyCode::Up => "Up".to_string(),
            KeyCode::Down => "Down".to_string(),
            _ => format!("{:?}", key.code),
        };
        state.log_event(&format!(
            "[KEY] {} | pane={} msg_idx={:?} scroll={}",
            key_str, pane_name, state.message_table_state.selected(), state.preview_scroll
        ));

        // --- Search bar input mode ---
        if state.search_bar_active {
            match key.code {
                KeyCode::Esc => {
                    state.search_bar_active = false;
                    if state.search_input.is_empty() {
                        state.search_mode = false;
                    }
                }
                KeyCode::Enter => {
                    state.search_bar_active = false;
                    if !state.search_input.is_empty() {
                        state.search_pending = true;
                        state.search_start = Some(Instant::now());
                    } else {
                        state.restore_all_messages();
                    }
                }
                KeyCode::Backspace => { state.search_input.pop(); }
                KeyCode::Char(c) => { state.search_input.push(c); }
                _ => {}
            }
            return Ok(());
        }

        match key.code {
            KeyCode::Char('q') => state.exit = true,
            KeyCode::Esc => {
                if state.search_mode {
                    state.restore_all_messages();
                } else {
                    state.exit = true;
                }
            }
            KeyCode::Char('/') => {
                state.search_bar_active = true;
            }
            KeyCode::Tab => {
                state.active_pane = match state.active_pane {
                    ActivePane::Messages => ActivePane::Preview,
                    ActivePane::Preview => ActivePane::Messages,
                };
            }
            KeyCode::Char('j') | KeyCode::Down => match state.active_pane {
                ActivePane::Messages => {
                    let next = state
                        .message_table_state
                        .selected()
                        .map(|i| (i + 1).min(state.message_row_ids.len().saturating_sub(1)))
                        .unwrap_or(0);
                    if !state.message_row_ids.is_empty() {
                        state.message_table_state.select(Some(next));
                        state.select_message(browser, next);
                    }
                }
                ActivePane::Preview => {
                    state.preview_scroll = state.preview_scroll.saturating_add(1);
                }
            },
            KeyCode::Char('k') | KeyCode::Up => match state.active_pane {
                ActivePane::Messages => {
                    if let Some(i) = state.message_table_state.selected()
                        && i > 0
                    {
                        state.message_table_state.select(Some(i - 1));
                        state.select_message(browser, i - 1);
                    }
                }
                ActivePane::Preview => {
                    state.preview_scroll = state.preview_scroll.saturating_sub(1);
                }
            },
            KeyCode::Enter => {
                if let Some(selected) = state.message_table_state.selected() {
                    state.select_message(browser, selected);
                    state.active_pane = ActivePane::Preview;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn main() {
    let cli = Cli::parse();

    match &cli.command {
        Commands::List { file, format, limit } => {
            if let Err(e) = list_emails(file, format.as_ref(), *limit) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
        Commands::Search { file, query, format } => {
            if let Err(e) = search_emails(file, query, format.as_ref()) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
        Commands::Browse { file, debug } => {
            if let Err(e) = browse_pst(file, *debug) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
        Commands::Stats { file } => {
            if let Err(e) = stats_pst(file) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
        Commands::Export { file, output } => {
            if let Err(e) = export_pst(file, output.as_ref()) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::rc::Rc;

    fn open_test_store(path: &str) -> (Rc<UnicodeStore>, Rc<UnicodeFolder>) {
        let pst = Rc::new(UnicodePstFile::open(path).unwrap());
        let store = UnicodeStore::read(Rc::clone(&pst)).unwrap();
        let entry_id = store.properties().ipm_sub_tree_entry_id().unwrap();
        let root = UnicodeFolder::read(Rc::clone(&store), &entry_id).unwrap();
        (store, root)
    }

    fn prop_to_string(v: &PropertyValue) -> Option<String> {
        match v {
            PropertyValue::String8(s) => Some(s.to_string()),
            PropertyValue::Unicode(s) => Some(s.to_string()),
            _ => None,
        }
    }

    // ── sample.pst tests ─────────────────────────────────────────────────────

    /// Verify the folder structure: 5 folders total, including the expected
    /// named folders from the Aspose sample file.
    #[test]
    fn test_sample_pst_folder_count() {
        let mut stats = PstStats::new();
        let (store, root) = open_test_store("testdata/sample.pst");
        collect_stats(Rc::clone(&store), &root, &mut stats);
        assert_eq!(stats.folder_count, 5);
    }

    /// Verify that exactly one email exists in the sample PST.
    #[test]
    fn test_sample_pst_email_count() {
        let mut stats = PstStats::new();
        let (store, root) = open_test_store("testdata/sample.pst");
        collect_stats(Rc::clone(&store), &root, &mut stats);
        assert_eq!(stats.email_count, 1);
        assert_eq!(stats.attachment_count, 0);
    }

    /// Verify that non-email artifact counts are all zero.
    #[test]
    fn test_sample_pst_no_other_artifacts() {
        let mut stats = PstStats::new();
        let (store, root) = open_test_store("testdata/sample.pst");
        collect_stats(Rc::clone(&store), &root, &mut stats);
        assert_eq!(stats.calendar_count, 0);
        assert_eq!(stats.contact_count, 0);
        assert_eq!(stats.task_count, 0);
        assert_eq!(stats.note_count, 0);
    }

    /// Verify the subject and sender of the single message in the sample.
    #[test]
    fn test_sample_pst_message_fields() {
        let (store, root) = open_test_store("testdata/sample.pst");

        // Walk folders until we find a message
        fn find_message(
            store: &Rc<UnicodeStore>,
            folder: &UnicodeFolder,
        ) -> Option<(String, String, String)> {
            if let Some(table) = folder.contents_table() {
                for row in table.rows_matrix() {
                    let entry_id = store
                        .properties()
                        .make_entry_id(NodeId::from(u32::from(row.id())))
                        .ok()?;
                    let msg = UnicodeMessage::read(
                        Rc::clone(store),
                        &entry_id,
                        Some(&[0x0037, 0x0C1A, 0x0E04]),
                    )
                    .ok()?;
                    let props = msg.properties();
                    let get = |id: u16| -> String {
                        props
                            .get(id)
                            .and_then(|v| match v {
                                PropertyValue::String8(s) => Some(s.to_string()),
                                PropertyValue::Unicode(s) => Some(s.to_string()),
                                _ => None,
                            })
                            .unwrap_or_default()
                    };
                    return Some((get(0x0037), get(0x0C1A), get(0x0E04)));
                }
            }
            if let Some(htable) = folder.hierarchy_table() {
                for row in htable.rows_matrix() {
                    let entry_id = store
                        .properties()
                        .make_entry_id(NodeId::from(u32::from(row.id())))
                        .ok()?;
                    if let Ok(sub) = UnicodeFolder::read(Rc::clone(store), &entry_id) {
                        if let Some(result) = find_message(store, &sub) {
                            return Some(result);
                        }
                    }
                }
            }
            None
        }

        let (subject, from, to) = find_message(&store, &root)
            .expect("expected at least one message in sample.pst");

        assert!(
            subject.contains("Aspose.Email"),
            "unexpected subject: {subject:?}"
        );
        assert_eq!(from, "Sender Name");
        assert!(to.contains("Recipient 1"), "unexpected To: {to:?}");
    }

    /// Verify the sample PST has no timestamps (the Aspose sample omits them).
    #[test]
    fn test_sample_pst_no_timestamps() {
        let mut stats = PstStats::new();
        let (store, root) = open_test_store("testdata/sample.pst");
        collect_stats(Rc::clone(&store), &root, &mut stats);
        assert!(
            stats.earliest_ts.is_none(),
            "expected no timestamps in sample.pst"
        );
        assert!(
            stats.latest_ts.is_none(),
            "expected no timestamps in sample.pst"
        );
    }

    // ── legacy test (requires testdata/outlook.pst) ───────────────────────────

    #[test]
    #[ignore = "requires testdata/outlook.pst which is not committed"]
    fn test_message_content_loads() {
        let (store, root) = open_test_store("testdata/outlook.pst");
        fn check_folder(store: &Rc<UnicodeStore>, folder: &UnicodeFolder, depth: usize) {
            let name = folder.properties().display_name().unwrap_or_default();
            if let Some(table) = folder.contents_table() {
                for row in table.rows_matrix() {
                    let entry_id = store.properties()
                        .make_entry_id(NodeId::from(u32::from(row.id()))).unwrap();
                    let props_filter = Some(&[0x0037u16, 0x0C1A, 0x0E04, 0x0E02, 0x0039, 0x0E06, 0x1000, 0x1013, 0x1009][..]);
                    match UnicodeMessage::read(Rc::clone(store), &entry_id, props_filter) {
                        Ok(msg) => {
                            let props = msg.properties();
                            let subj = props.get(0x0037).and_then(prop_to_string).unwrap_or("(none)".into());
                            let body_plain = props.get(0x1000).and_then(prop_to_string);
                            let body_html = props.get(0x1013).and_then(|v| match v {
                                PropertyValue::Binary(b) => Some(String::from_utf8_lossy(b.buffer()).into_owned()),
                                _ => prop_to_string(v),
                            });
                            eprintln!("{}{}/{}  plain={} html={}",
                                "  ".repeat(depth), name, subj,
                                body_plain.as_deref().map(|s| &s[..s.len().min(60)]).unwrap_or("NONE"),
                                body_html.as_deref().map(|s| &s[..s.len().min(60)]).unwrap_or("NONE"));
                        }
                        Err(e) => eprintln!("{}{}  ERROR: {}", "  ".repeat(depth), name, e),
                    }
                }
            }
            if let Some(htable) = folder.hierarchy_table() {
                for row in htable.rows_matrix() {
                    let entry_id = store.properties()
                        .make_entry_id(NodeId::from(u32::from(row.id()))).unwrap();
                    if let Ok(sub) = UnicodeFolder::read(Rc::clone(store), &entry_id) {
                        check_folder(store, &sub, depth + 1);
                    }
                }
            }
        }
        check_folder(&store, &root, 0);
    }

    // ── export tests ──────────────────────────────────────────────────────────

    #[test]
    fn test_export_sample_pst() {
        let db_path = std::env::temp_dir().join("pstexplorer_test_export.db");
        // Clean up from any previous run
        let _ = std::fs::remove_file(&db_path);

        export_pst(
            &PathBuf::from("testdata/sample.pst"),
            Some(&db_path),
        )
        .expect("export should succeed");

        let conn = Connection::open(&db_path).unwrap();

        // Verify folder count
        let folder_count: i64 = conn
            .query_row("SELECT COUNT(*) FROM folders", [], |r| r.get(0))
            .unwrap();
        assert_eq!(folder_count, 5);

        // Verify message count
        let msg_count: i64 = conn
            .query_row("SELECT COUNT(*) FROM messages", [], |r| r.get(0))
            .unwrap();
        assert_eq!(msg_count, 1);

        // Verify the message has the expected subject and sender
        let (subject, sender): (String, String) = conn
            .query_row(
                "SELECT subject, sender FROM messages LIMIT 1",
                [],
                |r| Ok((r.get(0)?, r.get(1)?)),
            )
            .unwrap();
        assert!(
            subject.contains("Aspose.Email"),
            "unexpected subject: {subject:?}"
        );
        assert_eq!(sender, "Sender Name");

        // Verify folder paths are populated
        let root_path: String = conn
            .query_row(
                "SELECT path FROM folders WHERE parent_id IS NULL",
                [],
                |r| r.get(0),
            )
            .unwrap();
        assert!(!root_path.is_empty());

        // Clean up
        let _ = std::fs::remove_file(&db_path);
    }
}
