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
use std::{io, path::PathBuf, rc::Rc};

/// Convert a Windows FILETIME (100-ns ticks since 1601-01-01 UTC) to a
/// human-readable UTC string, e.g. "2010-11-24 15:24:27 UTC".
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
    Folders,
    Messages,
    Preview,
}

struct AppState {
    exit: bool,
    current_folder: Rc<UnicodeFolder>,
    folder_list_state: ListState,
    message_list_state: ListState,
    folders: Vec<(String, usize, bool)>,
    messages: Vec<String>,
    current_message_content: String,
    current_headers: MessageHeaders,
    active_pane: ActivePane,
    preview_scroll: u16,
    /// The folder whose messages are currently shown in the message list.
    /// Differs from current_folder when hovering over a subfolder.
    messages_folder: Rc<UnicodeFolder>,
    /// Debug event log; None if debug mode not enabled.
    debug_log: Option<Vec<String>>,
    /// Transient status bar message (cleared on next keypress).
    status_message: Option<String>,
}

impl PstBrowser {
    fn new(store: Rc<UnicodeStore>, root_folder: Rc<UnicodeFolder>) -> Self {
        Self { store, root_folder }
    }
}

impl AppState {
    fn new(browser: &PstBrowser, debug: bool) -> Self {
        let folders = Self::get_folders(browser, &browser.root_folder);

        // Show messages from the first subfolder (if any), otherwise root folder.
        // Also keep track of which folder is being shown so select_message is correct.
        let messages_folder: Rc<UnicodeFolder> = if !folders.is_empty() {
            browser
                .root_folder
                .hierarchy_table()
                .and_then(|t| t.rows_matrix().next())
                .and_then(|row| {
                    let entry_id = browser
                        .store
                        .properties()
                        .make_entry_id(NodeId::from(u32::from(row.id())))
                        .ok()?;
                    UnicodeFolder::read(Rc::clone(&browser.store), &entry_id).ok()
                })
                .unwrap_or_else(|| Rc::clone(&browser.root_folder))
        } else {
            Rc::clone(&browser.root_folder)
        };
        let (messages, current_message_content) = Self::get_messages(browser, &messages_folder);

        let mut folder_list_state = ListState::default();
        if !folders.is_empty() {
            folder_list_state.select(Some(0));
        }

        Self {
            exit: false,
            current_folder: Rc::clone(&browser.root_folder),
            folder_list_state,
            message_list_state: ListState::default(),
            folders,
            messages,
            current_message_content,
            current_headers: MessageHeaders::default(),
            active_pane: ActivePane::Folders,
            preview_scroll: 0,
            messages_folder,
            debug_log: if debug { Some(vec![]) } else { None },
            status_message: None,
        }
    }

    fn log_event(&mut self, msg: &str) {
        if let Some(log) = &mut self.debug_log {
            log.push(msg.to_string());
        }
    }

    fn get_folders(browser: &PstBrowser, folder: &UnicodeFolder) -> Vec<(String, usize, bool)> {
        folder
            .hierarchy_table()
            .map(|table| {
                table
                    .rows_matrix()
                    .filter_map(|row| {
                        let entry_id = browser
                            .store
                            .properties()
                            .make_entry_id(NodeId::from(u32::from(row.id())))
                            .ok()?;
                        let subfolder =
                            UnicodeFolder::read(Rc::clone(&browser.store), &entry_id).ok()?;
                        let name = subfolder.properties().display_name().ok()?;
                        let count = subfolder
                            .contents_table()
                            .map(|t| t.rows_matrix().count())
                            .unwrap_or(0);
                        let has_subfolders = subfolder
                            .hierarchy_table()
                            .map(|t| t.rows_matrix().next().is_some())
                            .unwrap_or(false);
                        Some((name, count, has_subfolders))
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    fn preview_folder(&mut self, browser: &PstBrowser, index: usize) {
        let current_folder = Rc::clone(&self.current_folder);
        if let Some(table) = current_folder.hierarchy_table()
            && let Some(row) = table.rows_matrix().nth(index)
        {
            let entry_id = browser
                .store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
                .ok();
            if let Some(entry_id) = entry_id
                && let Ok(folder) = UnicodeFolder::read(Rc::clone(&browser.store), &entry_id)
            {
                let (messages, content) = Self::get_messages(browser, &folder);
                self.messages = messages;
                self.current_message_content = content;
                self.current_headers = MessageHeaders::default();
                self.message_list_state = ListState::default();
                self.preview_scroll = 0;
                self.messages_folder = folder;
            }
        }
    }

    fn get_messages(browser: &PstBrowser, folder: &UnicodeFolder) -> (Vec<String>, String) {
        let messages: Vec<String> = folder
            .contents_table()
            .map(|table| {
                table
                    .rows_matrix()
                    .map(|row| {
                        let entry_id = match browser
                            .store
                            .properties()
                            .make_entry_id(NodeId::from(u32::from(row.id())))
                        {
                            Ok(eid) => eid,
                            Err(e) => return format!("(entry id error: {})", e),
                        };
                        match UnicodeMessage::read(
                            Rc::clone(&browser.store),
                            &entry_id,
                            Some(&[0x0037]),
                        ) {
                            Ok(message) => message
                                .properties()
                                .get(0x0037)
                                .and_then(|v| match v {
                                    PropertyValue::String8(s) => Some(s.to_string()),
                                    PropertyValue::Unicode(s) => Some(s.to_string()),
                                    _ => None,
                                })
                                .unwrap_or_else(|| "(no subject)".to_string()),
                            Err(e) => format!("(unreadable: {})", e),
                        }
                    })
                    .collect()
            })
            .unwrap_or_default();

        let content = if messages.is_empty() {
            "No messages in this folder".to_string()
        } else {
            "Select a message to view its content".to_string()
        };

        (messages, content)
    }

    fn navigate_to_folder(&mut self, browser: &PstBrowser, index: usize) {
        // Clone current folder reference to avoid borrow issues
        let current_folder = Rc::clone(&self.current_folder);

        if let Some(table) = current_folder.hierarchy_table()
            && let Some(row) = table.rows_matrix().nth(index)
        {
            let entry_id = browser
                .store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
                .ok();

            if let Some(entry_id) = entry_id
                && let Ok(new_folder) = UnicodeFolder::read(Rc::clone(&browser.store), &entry_id)
            {
                self.current_folder = new_folder;
                self.folders = Self::get_folders(browser, &self.current_folder);
                self.current_headers = MessageHeaders::default();
                let mut folder_state = ListState::default();
                if !self.folders.is_empty() {
                    folder_state.select(Some(0));
                }
                self.folder_list_state = folder_state;
                self.message_list_state = ListState::default();
                self.preview_scroll = 0;
                self.active_pane = ActivePane::Folders;
                // Show first subfolder's messages if there are subfolders, else this folder's
                if !self.folders.is_empty() {
                    self.preview_folder(browser, 0);
                } else {
                    let (messages, content) = Self::get_messages(browser, &self.current_folder);
                    self.messages = messages;
                    self.current_message_content = content;
                    self.messages_folder = Rc::clone(&self.current_folder);
                }
            }
        }
    }

    fn select_message(&mut self, browser: &PstBrowser, index: usize) {
        let current_folder = Rc::clone(&self.messages_folder);

        if let Some(table) = current_folder.contents_table()
            && let Some(row) = table.rows_matrix().nth(index)
        {
            let entry_id = browser
                .store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
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
                            PropertyValue::Binary(b) => String::from_utf8(b.buffer().to_vec()).ok(),
                            PropertyValue::String8(s) => Some(s.to_string()),
                            PropertyValue::Unicode(s) => Some(s.to_string()),
                            _ => None,
                        })
                    })
                    .or_else(|| {
                        props.get(0x1009).and_then(|value| match value {
                            PropertyValue::Binary(b) => String::from_utf8(b.buffer().to_vec()).ok(),
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
                        if let Err(e) =
                            terminal.draw(|frame| draw_ui(frame, &browser, &mut app_state))
                        {
                            eprintln!("Error drawing UI: {}", e);
                            break;
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
            Constraint::Min(0),    // Main content
            Constraint::Length(1), // Status bar
        ])
        .split(frame.area());

    let main_layout = Layout::default()
        .direction(ratatui::layout::Direction::Horizontal)
        .constraints([
            Constraint::Percentage(30), // Folder list
            Constraint::Percentage(70), // Messages + preview
        ])
        .split(layout[0]);

    draw_folder_list(frame, state, main_layout[0]);
    draw_messages_pane(frame, state, main_layout[1]);

    let status_text = if let Some(msg) = &state.status_message {
        msg.clone()
    } else {
        match state.active_pane {
            ActivePane::Folders => " [Folders] j/k: navigate  Enter/l: open  h: back  Tab: → messages  D: dump state  q: quit".to_string(),
            ActivePane::Messages => " [Messages] j/k: navigate  Enter: view  Tab: → preview  D: dump state  q: quit".to_string(),
            ActivePane::Preview => " [Preview] j/k: scroll  Tab: → folders  D: dump state  q: quit".to_string(),
        }
    };
    let status_style = if state.status_message.is_some() {
        Style::default().fg(ratatui::style::Color::Green)
    } else {
        Style::default().fg(ratatui::style::Color::DarkGray)
    };
    let status = ratatui::widgets::Paragraph::new(status_text).style(status_style);
    frame.render_widget(status, layout[1]);
}

fn draw_folder_list(frame: &mut ratatui::Frame, state: &mut AppState, area: Rect) {
    let folder_name = state
        .current_folder
        .properties()
        .display_name()
        .unwrap_or_else(|_| "Root".to_string());

    let items: Vec<ListItem> = state
        .folders
        .iter()
        .map(|(name, count, has_sub)| {
            let label = if *has_sub && *count == 0 {
                format!("▶ {}", name)
            } else if *has_sub {
                format!("▶ {} ({})", name, count)
            } else {
                format!("{} ({})", name, count)
            };
            ListItem::new(label)
        })
        .collect();

    let border_style = if state.active_pane == ActivePane::Folders {
        Style::default().fg(ratatui::style::Color::Cyan)
    } else {
        Style::default()
    };

    let list = List::new(items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(border_style)
                .title(format!("Folders - {}", folder_name)),
        )
        .highlight_style(
            Style::default()
                .fg(ratatui::style::Color::Yellow)
                .add_modifier(ratatui::style::Modifier::BOLD),
        );

    frame.render_stateful_widget(list, area, &mut state.folder_list_state);
}

fn draw_messages_pane(frame: &mut ratatui::Frame, state: &mut AppState, area: Rect) {
    let layout = Layout::default()
        .direction(ratatui::layout::Direction::Vertical)
        .constraints([
            Constraint::Percentage(40), // Message list
            Constraint::Percentage(60), // Message preview
        ])
        .split(area);

    draw_message_list(frame, state, layout[0]);
    draw_message_preview(frame, state, layout[1]);
}

fn draw_message_list(frame: &mut ratatui::Frame, state: &mut AppState, area: Rect) {
    let items: Vec<ListItem> = state
        .messages
        .iter()
        .map(|subject| ListItem::new(subject.as_str()))
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
                .title("Messages"),
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
        // Clear transient status message on any keypress
        state.status_message = None;

        // Log keypress if debug mode enabled
        let pane_name = match state.active_pane {
            ActivePane::Folders => "Folders",
            ActivePane::Messages => "Messages",
            ActivePane::Preview => "Preview",
        };
        let folder_idx = state.folder_list_state.selected().unwrap_or(0);
        let msg_idx = state.message_list_state.selected().unwrap_or(0);
        let key_str = match key.code {
            KeyCode::Char(c) => format!("'{}'", c),
            KeyCode::Enter => "Enter".to_string(),
            KeyCode::Tab => "Tab".to_string(),
            KeyCode::Esc => "Esc".to_string(),
            KeyCode::Up => "Up".to_string(),
            KeyCode::Down => "Down".to_string(),
            KeyCode::Left => "Left".to_string(),
            KeyCode::Right => "Right".to_string(),
            _ => format!("{:?}", key.code),
        };
        state.log_event(&format!(
            "[KEY] {} | pane={} folder_idx={} msg_idx={} scroll={}",
            key_str, pane_name, folder_idx, msg_idx, state.preview_scroll
        ));

        match key.code {
            KeyCode::Char('q') | KeyCode::Esc => state.exit = true,
            KeyCode::Tab => {
                match state.active_pane {
                    ActivePane::Folders => {
                        state.active_pane = ActivePane::Messages;
                        if state.message_list_state.selected().is_none()
                            && !state.messages.is_empty()
                        {
                            state.message_list_state.select(Some(0));
                        }
                    }
                    ActivePane::Messages => {
                        state.active_pane = ActivePane::Preview;
                    }
                    ActivePane::Preview => {
                        state.active_pane = ActivePane::Folders;
                    }
                }
            }
            KeyCode::Char('j') | KeyCode::Down => match state.active_pane {
                ActivePane::Folders => {
                    let next = state
                        .folder_list_state
                        .selected()
                        .map(|i| (i + 1).min(state.folders.len().saturating_sub(1)))
                        .unwrap_or(0);
                    if !state.folders.is_empty() {
                        state.folder_list_state.select(Some(next));
                        state.preview_folder(browser, next);
                    }
                }
                ActivePane::Messages => {
                    let next = state
                        .message_list_state
                        .selected()
                        .map(|i| (i + 1).min(state.messages.len().saturating_sub(1)))
                        .unwrap_or(0);
                    if !state.messages.is_empty() {
                        state.message_list_state.select(Some(next));
                        state.select_message(browser, next);
                    }
                }
                ActivePane::Preview => {
                    state.preview_scroll = state.preview_scroll.saturating_add(1);
                }
            },
            KeyCode::Char('k') | KeyCode::Up => match state.active_pane {
                ActivePane::Folders => {
                    if let Some(i) = state.folder_list_state.selected()
                        && i > 0
                    {
                        state.folder_list_state.select(Some(i - 1));
                        state.preview_folder(browser, i - 1);
                    }
                }
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
            KeyCode::Char('l') | KeyCode::Right | KeyCode::Enter => {
                match state.active_pane {
                    ActivePane::Folders => {
                        if let Some(selected) = state.folder_list_state.selected() {
                            state.navigate_to_folder(browser, selected);
                        }
                    }
                    ActivePane::Messages => {
                        if let Some(selected) = state.message_list_state.selected() {
                            state.select_message(browser, selected);
                            state.active_pane = ActivePane::Preview;
                        }
                    }
                    ActivePane::Preview => {}
                }
            }
            KeyCode::Char('D') => {
                let folder_name = state
                    .current_folder
                    .properties()
                    .display_name()
                    .unwrap_or_else(|_| "?".to_string());
                let messages_folder_name = state
                    .messages_folder
                    .properties()
                    .display_name()
                    .unwrap_or_else(|_| "?".to_string());
                let raw_row_count = state
                    .messages_folder
                    .contents_table()
                    .map(|t| t.rows_matrix().count())
                    .unwrap_or(0);
                let pane = match state.active_pane {
                    ActivePane::Folders => "Folders",
                    ActivePane::Messages => "Messages",
                    ActivePane::Preview => "Preview",
                };
                let dump = format!(
                    "=== pstexplorer state dump ===\n\
                     active_pane:       {}\n\
                     current_folder:    {} ({} subfolders)\n\
                     messages_folder:   {} (raw rows={} displayed={})\n\
                     folder_idx:        {:?}\n\
                     message_idx:       {:?}\n\
                     preview_scroll:    {}\n\
                     debug_mode:        {}\n",
                    pane,
                    folder_name,
                    state.folders.len(),
                    messages_folder_name,
                    raw_row_count,
                    state.messages.len(),
                    state.folder_list_state.selected(),
                    state.message_list_state.selected(),
                    state.preview_scroll,
                    state.debug_log.is_some(),
                );
                let dump_path = "pstexplorer-state-dump.txt";
                match std::fs::write(dump_path, &dump) {
                    Ok(_) => {
                        state.log_event(&format!("[DUMP] State dumped to {}", dump_path));
                        state.status_message = Some(format!(" State dumped to {}", dump_path));
                    }
                    Err(e) => {
                        state.status_message = Some(format!(" Dump failed: {}", e));
                    }
                }
            }
            KeyCode::Char('h') | KeyCode::Left => {
                state.current_folder = Rc::clone(&browser.root_folder);
                state.folders = AppState::get_folders(browser, &state.current_folder);
                state.current_headers = MessageHeaders::default();
                let mut folder_state = ListState::default();
                if !state.folders.is_empty() {
                    folder_state.select(Some(0));
                }
                state.folder_list_state = folder_state;
                state.message_list_state = ListState::default();
                state.preview_scroll = 0;
                state.active_pane = ActivePane::Folders;
                if !state.folders.is_empty() {
                    state.preview_folder(browser, 0);
                } else {
                    let (messages, content) =
                        AppState::get_messages(browser, &state.current_folder);
                    state.messages = messages;
                    state.current_message_content = content;
                    state.messages_folder = Rc::clone(&state.current_folder);
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
                                PropertyValue::Binary(b) => String::from_utf8(b.buffer().to_vec()).ok(),
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
