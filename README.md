# pstexplorer

A CLI tool to explore and extract data from Outlook PST files.

## Usage

```
pstexplorer <COMMAND>

Commands:
  list    List all emails in a PST file
  search  Search emails in a PST file by query string (matches from, to, cc, body)
  browse  Browse PST file contents in a TUI
  stats   Print statistics about a PST file
  export  Export a PST file to a SQLite database
```

## Features

### list

List all emails with subject, sender, recipient, and date. Supports `--format csv|tsv|json` for structured output and `--limit` to cap the number of entries.

### search

Case-insensitive full-text search across from, to, cc, and body fields. Supports the same `--format` options as `list`.

### browse

Interactive terminal UI for navigating folders and reading messages.

### stats

Print a summary of the PST file: folder count, email/calendar/contact/task/note counts, attachment count, and date range.

### export

Export folders and messages to a SQLite database for further analysis. Use `--output` to set the database path and `--limit` to cap the number of exported messages.
