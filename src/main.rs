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
    style::Style,
    widgets::{Block, Borders, List, ListItem, ListState, Paragraph},
};
use std::{io, path::PathBuf, rc::Rc};

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

#[derive(PartialEq)]
enum ActivePane {
    Folders,
    Messages,
}

struct AppState {
    exit: bool,
    current_folder: Rc<UnicodeFolder>,
    folder_list_state: ListState,
    message_list_state: ListState,
    folders: Vec<String>,
    messages: Vec<String>,
    current_message_content: String,
    active_pane: ActivePane,
}

impl PstBrowser {
    fn new(store: Rc<UnicodeStore>, root_folder: Rc<UnicodeFolder>) -> Self {
        Self { store, root_folder }
    }
}

impl AppState {
    fn new(browser: &PstBrowser) -> Self {
        let folders = Self::get_folders(browser, &browser.root_folder);
        let (messages, current_message_content) = Self::get_messages(browser, &browser.root_folder);

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
            active_pane: ActivePane::Folders,
        }
    }

    fn get_folders(browser: &PstBrowser, folder: &UnicodeFolder) -> Vec<String> {
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
                        subfolder.properties().display_name().ok()
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    fn get_messages(browser: &PstBrowser, folder: &UnicodeFolder) -> (Vec<String>, String) {
        let messages: Vec<String> = folder
            .contents_table()
            .map(|table| {
                table
                    .rows_matrix()
                    .filter_map(|row| {
                        let entry_id = browser
                            .store
                            .properties()
                            .make_entry_id(NodeId::from(u32::from(row.id())))
                            .ok()?;
                        let message = UnicodeMessage::read(
                            Rc::clone(&browser.store),
                            &entry_id,
                            Some(&[0x0037]),
                        )
                        .ok()?;
                        Some(
                            message
                                .properties()
                                .get(0x0037)
                                .and_then(|v| match v {
                                    PropertyValue::String8(s) => Some(s.to_string()),
                                    PropertyValue::Unicode(s) => Some(s.to_string()),
                                    _ => None,
                                })
                                .unwrap_or_else(|| "(no subject)".to_string()),
                        )
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
                let (messages, content) = Self::get_messages(browser, &self.current_folder);
                self.messages = messages;
                self.current_message_content = content;
                let mut folder_state = ListState::default();
                if !self.folders.is_empty() {
                    folder_state.select(Some(0));
                }
                self.folder_list_state = folder_state;
                self.message_list_state = ListState::default();
                self.active_pane = ActivePane::Folders;
            }
        }
    }

    fn select_message(&mut self, browser: &PstBrowser, index: usize) {
        // Clone current folder reference to avoid borrow issues
        let current_folder = Rc::clone(&self.current_folder);

        if let Some(table) = current_folder.contents_table()
            && let Some(row) = table.rows_matrix().nth(index)
        {
            let entry_id = browser
                .store
                .properties()
                .make_entry_id(NodeId::from(u32::from(row.id())))
                .ok();

            if let Some(entry_id) = entry_id
                && let Ok(message) =
                    UnicodeMessage::read(Rc::clone(&browser.store), &entry_id, None)
            {
                self.current_message_content = message
                    .properties()
                    .get(0x1000)
                    .and_then(|value| match value {
                        PropertyValue::String8(s) => Some(s.to_string()),
                        PropertyValue::Unicode(s) => Some(s.to_string()),
                        _ => None,
                    })
                    .or_else(|| {
                        message
                            .properties()
                            .get(0x1013)
                            .and_then(|value| match value {
                                PropertyValue::Binary(b) => {
                                    String::from_utf8(b.buffer().to_vec()).ok()
                                }
                                PropertyValue::String8(s) => Some(s.to_string()),
                                PropertyValue::Unicode(s) => Some(s.to_string()),
                                _ => None,
                            })
                    })
                    .or_else(|| {
                        message
                            .properties()
                            .get(0x1009)
                            .and_then(|value| match value {
                                PropertyValue::Binary(b) => {
                                    String::from_utf8(b.buffer().to_vec()).ok()
                                }
                                _ => None,
                            })
                    })
                    .unwrap_or_else(|| "No message content available".to_string());
            }
        }
    }
}

fn browse_pst(file_path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
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
                    let mut app_state = AppState::new(&browser);

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

    let help = match state.active_pane {
        ActivePane::Folders => " [Folders] j/k: navigate  Enter/l: open  h: back  Tab: switch to messages  q: quit",
        ActivePane::Messages => " [Messages] j/k: navigate  Enter: view  Tab: switch to folders  q: quit",
    };
    let status = ratatui::widgets::Paragraph::new(help)
        .style(Style::default().fg(ratatui::style::Color::DarkGray));
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
        .map(|name| ListItem::new(name.as_str()))
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
    let preview = Paragraph::new(state.current_message_content.as_str())
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title("Message Preview"),
        )
        .wrap(ratatui::widgets::Wrap { trim: true });

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
            },
            KeyCode::Char('k') | KeyCode::Up => match state.active_pane {
                ActivePane::Folders => {
                    if let Some(i) = state.folder_list_state.selected()
                        && i > 0
                    {
                        state.folder_list_state.select(Some(i - 1));
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
                        }
                    }
                }
            }
            KeyCode::Char('h') | KeyCode::Left => {
                let root_folder = Rc::clone(&browser.root_folder);
                state.current_folder = Rc::clone(&root_folder);
                state.folders = AppState::get_folders(browser, &state.current_folder);
                let (messages, content) = AppState::get_messages(browser, &state.current_folder);
                state.messages = messages;
                state.current_message_content = content;
                let mut folder_state = ListState::default();
                if !state.folders.is_empty() {
                    folder_state.select(Some(0));
                }
                state.folder_list_state = folder_state;
                state.message_list_state = ListState::default();
                state.active_pane = ActivePane::Folders;
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
        Commands::Browse { file } => {
            if let Err(e) = browse_pst(file) {
                eprintln!("Error: {}", e);
                std::process::exit(1);
            }
        }
    }
}
