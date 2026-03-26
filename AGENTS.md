# pstexplorer — Agent Reference

## Project overview

`pstexplorer` is a Rust CLI tool for reading, searching, and exporting data from Outlook PST (Personal Storage Table) files. It is read-only; it never writes back to the PST.

Repository: https://github.com/stchris/pstexplorer

## Stack

| Layer | Choice |
|---|---|
| Language | Rust (edition 2024) |
| PST parsing | `outlook-pst 1.1` (read-only, Unicode variant only) |
| CLI | `clap 4.5` with derive macros + `env` feature |
| TUI | `ratatui 0.30` + `crossterm 0.29` |
| SQLite export | `rusqlite 0.31` (bundled feature) |
| HTTP client | `ureq 3` (synchronous, `json` feature) — used for ChromaDB and embedding APIs |
| Serialization | `serde 1` + `serde_json 1` |
| FTM export | `ftm-types 0.4` |
| RTF handling | `compressed-rtf 1` + `rtf-parser 0.4` |
| Date/time | `chrono 0.4` |

The project is a single Rust crate with all source code in `src/main.rs`. There is no async runtime.

## PST reading conventions

- All PST operations use the Unicode variants: `UnicodePstFile`, `UnicodeStore`, `UnicodeFolder`, `UnicodeMessage`.
- The standard open sequence is:
  ```rust
  let pst = UnicodePstFile::open(path)?;
  let store = Rc::new(UnicodeStore::read(Rc::new(pst))?);
  let root_entry_id = store.properties().ipm_sub_tree_entry_id()?;
  let root_folder = UnicodeFolder::read(Rc::clone(&store), &root_entry_id)?;
  ```
- Messages are accessed via `folder.contents_table()` → `rows_matrix()`.
- Subfolders are accessed via `folder.hierarchy_table()` → `rows_matrix()`.
- Properties are read by hex ID, e.g. `props.get(0x0037)` for subject.

Key MAPI property IDs used:

| Hex ID | Name | Field |
|---|---|---|
| 0x0037 | PR_SUBJECT | Subject |
| 0x001A | PR_MESSAGE_CLASS | Item type (email/calendar/contact/task/note) |
| 0x0039 | PR_CLIENT_SUBMIT_TIME | Submit timestamp |
| 0x0C1A | PR_SENDER_NAME | From |
| 0x0E02 | PR_CC_RECIPIENTS | CC |
| 0x0E04 | PR_TO_RECIPIENTS | To |
| 0x0E06 | PR_MESSAGE_DELIVERY_TIME | Delivery timestamp |
| 0x0E13 | PR_ATTACH_NUM | Attachment count |
| 0x1000 | PR_BODY | Plain-text body |
| 0x1009 | PR_RTF_COMPRESSED | Compressed RTF body |
| 0x1013 | PR_HTML | HTML body |

Body extraction priority: plain text → HTML (via `html_to_text()`) → compressed RTF (via `rtf_compressed_to_text()`).

Timestamps are Windows FILETIME (100-ns ticks since 1601-01-01 UTC). Helpers: `filetime_to_string()` (display), `filetime_to_iso()` (ISO 8601).

## Core data structures

```rust
struct EmailRecord {
    id: String,
    folder: String,
    subject: String,
    from: String,
    to: String,
    cc: String,
    date: String,       // ISO 8601
    body_text: Option<String>,
    body_html: Option<String>,
}
```

Recursive collection helper: `collect_emails(store, folder, records, include_body)` — reused by `list`, `search`, and `llm embed` commands.

## Commands

### `list`
Lists all emails with optional `--format` (csv / tsv / json / ftm / text) and `--limit`.

### `search`
Case-insensitive full-text search across from, to, cc, and body. Same output formats as `list`.

### `browse`
Interactive ratatui TUI with two-pane layout (message list + preview). Features: `j`/`k` navigation, `/` search, `s` sort popup, `Tab`/`Enter` pane switch, `q` quit. Lazy-loads message rows for the visible window.

### `stats`
Counts folders, emails, attachments, calendar items, contacts, tasks, and notes. Reports earliest/latest timestamps. Item type is determined from PR_MESSAGE_CLASS (0x001A).

### `export`
Exports to a SQLite database (default: `<pst-stem>.db`). Schema: `folders`, `messages`, `attachments` tables with indexes. Uses WAL mode and a single transaction.

### `llm` (nested subcommand)

The `llm` command groups two subcommands for RAG over emails. Implemented with a nested `LlmCommands` enum in the `Commands` enum. HTTP calls use `ureq` (synchronous). ChromaDB's REST API **requires** embeddings to be supplied — it cannot generate them server-side.

#### `llm embed`

Reads all emails from a PST, generates embeddings via an OpenAI-compatible API, and upserts documents + metadata into a ChromaDB collection.

- Collection name defaults to the PST filename stem (e.g. `archive.pst` → `archive`).
- Documents are composed as `"Subject: {subject}\n\n{body}"`.
- Metadata stored per document: `folder`, `subject`, `from`, `to`, `cc`, `date`, `pst_id`.
- Document IDs in ChromaDB: `pst-{record.id}` (idempotent — safe to re-run).
- Implemented in `embed_emails()`. Uses `collect_emails(..., include_body=true)` then calls `call_embeddings_api()` per batch and `chroma_add_documents()`.

| Flag | Default | Notes |
|---|---|---|
| `--chroma-url` | `http://localhost:8000` | ChromaDB server |
| `--collection` | PST filename stem | Collection name |
| `--embedding-url` | _(required)_ | OpenAI-compatible embeddings base URL |
| `--embedding-key` | — | API key; also `EMBEDDING_API_KEY` env var |
| `--embedding-model` | `text-embedding-3-small` | Model name |
| `--batch-size` | 100 | Documents per ChromaDB add request |
| `--limit` | 0 (no limit) | Cap on messages processed |
| `--tenant` / `--database` | `default_tenant` / `default_database` | ChromaDB tenant/database |

#### `llm ask`

Takes a natural language question, embeds it, queries ChromaDB for the top-k most relevant emails, and sends them as context to a chat model.

- `--collection` is required (no PST file to derive a stem from).
- The embedding model **must match** the one used during `llm embed`.
- Implemented in `ask_llm()`. Uses `call_embeddings_api()` → `chroma_query()` → `call_chat_api()`.
- Context block format per result: `[N] Folder: … | From: … | Date: … | Subject: …\n\n{document}`.
- Chat API uses OpenAI-compatible `/chat/completions` endpoint.

| Flag | Default | Notes |
|---|---|---|
| `--collection` | _(required)_ | ChromaDB collection to query |
| `--chroma-url` | `http://localhost:8000` | ChromaDB server |
| `--n-results` / `-n` | 5 | Number of emails retrieved as context |
| `--embedding-url` | _(required)_ | OpenAI-compatible embeddings base URL |
| `--embedding-key` | — | API key; also `EMBEDDING_API_KEY` env var |
| `--embedding-model` | `text-embedding-3-small` | Must match model used at embed time |
| `--llm-url` | _(required)_ | OpenAI-compatible chat completions base URL |
| `--llm-key` | — | API key; also `LLM_API_KEY` env var |
| `--llm-model` | `gpt-4o-mini` | Chat model name |
| `--tenant` / `--database` | `default_tenant` / `default_database` | ChromaDB tenant/database |

#### Key helper functions (ChromaDB REST API)

| Function | Purpose |
|---|---|
| `chroma_heartbeat(url)` | `GET /api/v2/heartbeat` — validates server is reachable |
| `chroma_get_or_create_collection(...)` | GET collection, POST to create if 404 |
| `chroma_add_documents(...)` | `POST .../collections/{id}/add` with ids/documents/metadatas/embeddings |
| `chroma_query(...)` | `POST .../collections/{id}/query` — returns documents + metadatas arrays |
| `chroma_post(url, body)` | Shared helper: POST with error body surfaced on non-2xx |
| `call_embeddings_api(url, key, model, texts)` | OpenAI-compatible `/embeddings` → `Vec<Vec<f32>>` |
| `call_chat_api(url, key, model, system, user)` | OpenAI-compatible `/chat/completions` → answer string |

### `llmquery/query.py`

A standalone Python alternative to `llm ask`, using the `chromadb` and `ollama` Python libraries directly. Requires a locally running Ollama instance. Uses inline script metadata (PEP 723) so only `uv` is needed — no virtual environment setup.

```bash
./llmquery/query.py "who sent me invoices in 2023?"
./llmquery/query.py --collection testPST "what meetings did I have?"
./llmquery/query.py --collection testPST --n-results 10 "any emails about the budget?"
```

Models are hardcoded at the top of the file (`MODEL_EMBED`, `MODEL_CHAT`). ChromaDB host/port are also hardcoded.

## Output formats

`list` and `search` support `-–format`:
- `text` (default) — human-readable
- `csv` / `tsv` — delimited, with header row
- `json` — pretty-printed array
- `ftm` — FunTimeMetadata entity format

## Test data

- `testdata/sample.pst` — 1 email
- `testdata/testPST.pst` — ~6 emails

## Tooling

- Build: `cargo build`
- Run: `cargo run -- <command> [args]`
- Python fixtures: `uv run --with faker --with requests generate.py`
- Use `uv`/`uvx` for any Python tooling; never `pip` or `python` directly.
