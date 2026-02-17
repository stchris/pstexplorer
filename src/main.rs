use clap::{Parser, Subcommand};
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
    widgets::{Block, Borders, List, ListItem, ListState, Paragraph},
};
use chrono::{TimeZone, Utc};
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
    },
    /// Search emails in a PST file by query string (matches from, to, cc, body)
    Search {
        /// Path to the PST file
        #[arg(required = true)]
        file: PathBuf,
        /// Search query (case-insensitive, matched against from, to, cc, and body)
        #[arg(required = true)]
        query: String,
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
}

fn list_emails(file_path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    println!("Listing emails from: {:?}", file_path);

    // Open the PST file
    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);

    // Get the root folder hierarchy
    let hierarchy_table = store.root_hierarchy_table()?;

    println!("PST File Information:");
    println!(
        "  Number of rows in hierarchy: {}",
        hierarchy_table.rows_matrix().count()
    );

    // Get the IPM subtree (where emails are stored)
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = outlook_pst::messaging::folder::UnicodeFolder::read(
        Rc::clone(&store),
        &ipm_sub_tree_entry_id,
    )?;

    // Traverse the folder hierarchy and extract email information
    let total_emails = traverse_folder_hierarchy(Rc::clone(&store), &ipm_subtree_folder)?;

    println!("\nFound {} emails in the PST file", total_emails);

    Ok(())
}

fn search_emails(file_path: &PathBuf, query: &str) -> Result<(), Box<dyn std::error::Error>> {
    let query_lower = query.to_ascii_lowercase();
    let pst = UnicodePstFile::open(file_path)?;
    let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);
    let ipm_sub_tree_entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let ipm_subtree_folder = outlook_pst::messaging::folder::UnicodeFolder::read(
        Rc::clone(&store),
        &ipm_sub_tree_entry_id,
    )?;
    let total = search_traverse_folders(Rc::clone(&store), &ipm_subtree_folder, &query_lower)?;
    println!("\nFound {} matching emails", total);
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
                // Subject, From, To, CC, Received Time, Plain body, HTML body, RTF body
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

                // Resolve body: plain text, then HTML, then RTF
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
                Some(&[0x0037, 0x0C1A, 0x0E04, 0x0E02, 0x1000, 0x1013, 0x1009]),
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
                results.push(SearchResultItem {
                    folder_name: folder_name.clone(),
                    subject,
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
            }
        }
    }

    // Recursively traverse subfolders
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
            email_count += traverse_folder_hierarchy(store.clone(), &subfolder)?;
        }
    }

    Ok(email_count)
}

// TUI Data Structures
struct PstBrowser {
    store: Rc<UnicodeStore>,
    root_folder: Rc<UnicodeFolder>,
}

struct SearchResultItem {
    folder_name: String,
    subject: String,
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
    /// Lazily loaded subjects; None = not yet fetched.
    message_subjects: Vec<Option<String>>,
    message_list_state: ListState,
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

        let mut message_list_state = ListState::default();
        if n > 0 {
            message_list_state.select(Some(0));
        }

        Self {
            exit: false,
            message_row_ids: all_row_ids.clone(),
            message_folder_names: all_folder_names.clone(),
            all_row_ids,
            all_folder_names,
            message_subjects: vec![None; n],
            message_list_state,
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

    /// Load subjects for the visible window around the current scroll offset.
    fn load_visible_subjects(&mut self, browser: &PstBrowser) {
        let offset = self.message_list_state.offset();
        let end = (offset + self.message_list_height + 5).min(self.message_row_ids.len());
        for i in offset..end {
            if self.message_subjects[i].is_none() {
                let subject = browser
                    .store
                    .properties()
                    .make_entry_id(NodeId::from(self.message_row_ids[i]))
                    .ok()
                    .and_then(|eid| {
                        UnicodeMessage::read(Rc::clone(&browser.store), &eid, Some(&[0x0037])).ok()
                    })
                    .and_then(|msg| {
                        msg.properties().get(0x0037).and_then(|v| match v {
                            PropertyValue::String8(s) => Some(s.to_string()),
                            PropertyValue::Unicode(s) => Some(s.to_string()),
                            _ => None,
                        })
                    })
                    .unwrap_or_else(|| "(no subject)".to_string());
                self.message_subjects[i] = Some(subject);
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
        self.message_subjects = vec![None; n];
        self.search_mode = false;
        self.search_input.clear();
        self.message_list_state = ListState::default();
        if n > 0 { self.message_list_state.select(Some(0)); }
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
                        // Load subjects for the currently visible window before drawing.
                        app_state.load_visible_subjects(&browser);

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
                            app_state.message_subjects = results.iter().map(|r| Some(r.subject.clone())).collect();
                            app_state.message_list_state = ListState::default();
                            if n > 0 {
                                app_state.message_list_state.select(Some(0));
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
            list_emails(file_path)?;
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
    // Record visible height so load_visible_subjects knows the window size.
    // Subtract 2 for the border.
    state.message_list_height = area.height.saturating_sub(2) as usize;

    let count = state.message_row_ids.len();
    let selected_num = state.message_list_state.selected().map(|i| i + 1).unwrap_or(0);
    let title = if state.search_mode {
        format!("Search Results ({}/{})", selected_num, count)
    } else {
        format!("Messages ({}/{})", selected_num, count)
    };

    let items: Vec<ListItem> = (0..count)
        .map(|i| {
            let folder = state.message_folder_names.get(i).map(|s| s.as_str()).unwrap_or("?");
            let subject = state.message_subjects[i].as_deref().unwrap_or("…");
            ListItem::new(format!("{} | {}", folder, subject))
        })
        .collect();

    let border_style = if state.active_pane == ActivePane::Messages {
        Style::default().fg(ratatui::style::Color::Cyan)
    } else {
        Style::default()
    };

    let list = List::new(items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(border_style)
                .title(title),
        )
        .highlight_style(
            Style::default()
                .fg(ratatui::style::Color::Green)
                .add_modifier(ratatui::style::Modifier::BOLD),
        );

    frame.render_stateful_widget(list, area, &mut state.message_list_state);
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
            key_str, pane_name, state.message_list_state.selected(), state.preview_scroll
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
                        .message_list_state
                        .selected()
                        .map(|i| (i + 1).min(state.message_row_ids.len().saturating_sub(1)))
                        .unwrap_or(0);
                    if !state.message_row_ids.is_empty() {
                        state.message_list_state.select(Some(next));
                        state.select_message(browser, next);
                    }
                }
                ActivePane::Preview => {
                    state.preview_scroll = state.preview_scroll.saturating_add(1);
                }
            },
            KeyCode::Char('k') | KeyCode::Up => match state.active_pane {
                ActivePane::Messages => {
                    if let Some(i) = state.message_list_state.selected()
                        && i > 0
                    {
                        state.message_list_state.select(Some(i - 1));
                        state.select_message(browser, i - 1);
                    }
                }
                ActivePane::Preview => {
                    state.preview_scroll = state.preview_scroll.saturating_sub(1);
                }
            },
            KeyCode::Enter => {
                if let Some(selected) = state.message_list_state.selected() {
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
        Commands::List { file } => {
            if let Err(e) = list_emails(file) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
        Commands::Search { file, query } => {
            if let Err(e) = search_emails(file, query) {
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

    #[test]
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
}
