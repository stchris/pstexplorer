// PSTBrowser.swift
// Single-file SwiftUI email browser for pstexplorer SQLite exports.
//
// Build & run from the command line:
//   swiftc -parse-as-library -o PSTBrowser PSTBrowser.swift \
//          -framework SwiftUI -framework AppKit -framework WebKit \
//          -Xlinker -lsqlite3
//   ./PSTBrowser
//
// Or in Xcode: create a new macOS App project (SwiftUI, minimum macOS 14),
// replace ContentView.swift with this file, and add libsqlite3.tbd to
// "Frameworks, Libraries, and Embedded Content".

import SwiftUI
import AppKit
import WebKit
import SQLite3

// MARK: - Models

struct Folder: Identifiable, Hashable {
    let id: Int64
    let parentId: Int64?
    let name: String
    let path: String
}

struct Message: Identifiable, Hashable {
    let id: Int64
    let folderId: Int64
    let subject: String
    let sender: String
    let toRecipients: String
    let ccRecipients: String
    let submitTime: Date?
    let deliveryTime: Date?
    let bodyText: String?
    let bodyHtml: String?
    let attachmentCount: Int

    static func == (lhs: Message, rhs: Message) -> Bool { lhs.id == rhs.id }
    func hash(into hasher: inout Hasher) { hasher.combine(id) }
}

struct Attachment: Identifiable {
    let id: Int64
    let messageId: Int64
    let filename: String
    let contentType: String
    let size: Int
    let data: Data?
}

// MARK: - Database

final class Database: ObservableObject {
    @Published var folders: [Folder] = []
    @Published var messages: [Message] = []
    @Published var hasMore = false
    @Published var error: String?

    private var db: OpaquePointer?
    private let pageSize = 200

    // Stored so loadNextPage() can fetch the next page without re-specifying params.
    private var lastFolderId:  Int64?
    private var lastSearch    = ""
    private var lastSort      = SortField.deliveryTime
    private var lastAscending = false
    private var lastFromDate:  Date? = nil
    private var lastToDate:    Date? = nil

    private static let sqlDateFormatter: DateFormatter = {
        let f = DateFormatter()
        f.dateFormat = "yyyy-MM-dd"
        f.timeZone = TimeZone(identifier: "UTC")
        return f
    }()

    var isOpen: Bool { db != nil }

    deinit { close() }

    private func log(_ message: String) {
        print("[PSTBrowser] \(message)")
    }

    private func sqliteError() -> String {
        guard let db else { return "no connection" }
        return String(cString: sqlite3_errmsg(db))
    }

    func open(url: URL) {
        close()
        // immutable=1: skip WAL/SHM files — safe for a read-only export viewer.
        let uri = url.absoluteString + "?immutable=1"
        log("Opening \(uri)")
        let rc = sqlite3_open_v2(uri, &db, SQLITE_OPEN_READONLY | SQLITE_OPEN_URI, nil)
        guard rc == SQLITE_OK else {
            let msg = "Cannot open database: \(url.lastPathComponent) — \(sqliteError()) (rc=\(rc))"
            log(msg)
            error = msg
            return
        }
        log("Opened successfully")
        loadFolders()
        loadMessages(folderId: nil)
    }

    func close() {
        if db != nil { sqlite3_close(db); db = nil }
    }

    func loadFolders() {
        guard let db else { log("loadFolders: no open db"); return }
        let sql = "SELECT id, parent_id, name, path FROM folders ORDER BY path"
        var stmt: OpaquePointer?
        guard sqlite3_prepare_v2(db, sql, -1, &stmt, nil) == SQLITE_OK else {
            log("loadFolders prepare failed: \(sqliteError())")
            return
        }
        defer { sqlite3_finalize(stmt) }
        var result: [Folder] = []
        while sqlite3_step(stmt) == SQLITE_ROW {
            let id    = sqlite3_column_int64(stmt, 0)
            let pid   = sqlite3_column_type(stmt, 1) == SQLITE_NULL ? nil : Optional(sqlite3_column_int64(stmt, 1))
            let name  = String(cString: sqlite3_column_text(stmt, 2))
            let path  = String(cString: sqlite3_column_text(stmt, 3))
            result.append(Folder(id: id, parentId: pid, name: name, path: path))
        }
        log("loadFolders: loaded \(result.count) folder(s)")
        DispatchQueue.main.async { self.folders = result }
    }

    private struct ParsedSearch {
        var from: String?
        var to: String?
        var cc: String?
        var subject: String?
        var id: Int64?
        var hasAttachment = false
        var freeText: String?
    }

    private func parseSearch(_ raw: String) -> ParsedSearch {
        var p = ParsedSearch()
        var remaining: [String] = []
        for token in raw.split(separator: " ").map(String.init) {
            let lower = token.lowercased()
            if lower.hasPrefix("from:"),        let v = tokenValue(token, prefix: 5) { p.from = v }
            else if lower.hasPrefix("to:"),     let v = tokenValue(token, prefix: 3) { p.to = v }
            else if lower.hasPrefix("cc:"),     let v = tokenValue(token, prefix: 3) { p.cc = v }
            else if lower.hasPrefix("subject:"), let v = tokenValue(token, prefix: 8) { p.subject = v }
            else if lower.hasPrefix("id:"),     let v = tokenValue(token, prefix: 3) { p.id = Int64(v) }
            else if lower == "has:attachment" { p.hasAttachment = true }
            else { remaining.append(token) }
        }
        let ft = remaining.joined(separator: " ")
        p.freeText = ft.isEmpty ? nil : ft
        return p
    }

    private func tokenValue(_ token: String, prefix: Int) -> String? {
        let v = String(token.dropFirst(prefix))
        return v.isEmpty ? nil : v
    }

    func loadMessages(folderId: Int64?, search: String = "", sort: SortField = .deliveryTime,
                      ascending: Bool = false, fromDate: Date? = nil, toDate: Date? = nil) {
        lastFolderId  = folderId
        lastSearch    = search
        lastSort      = sort
        lastAscending = ascending
        lastFromDate  = fromDate
        lastToDate    = toDate
        execute(offset: 0, appending: false)
    }

    func loadNextPage() {
        guard hasMore else { return }
        execute(offset: messages.count, appending: true)
    }

    private func execute(offset: Int, appending: Bool) {
        guard let db else { return }
        var conditions: [String] = []
        var bindings: [Any] = []

        if let fid = lastFolderId {
            conditions.append("folder_id = ?")
            bindings.append(fid)
        }
        let parsed = parseSearch(lastSearch)
        if let v = parsed.from     { conditions.append("sender LIKE ?");        bindings.append("%\(v)%") }
        if let v = parsed.to       { conditions.append("to_recipients LIKE ?"); bindings.append("%\(v)%") }
        if let v = parsed.cc       { conditions.append("cc_recipients LIKE ?"); bindings.append("%\(v)%") }
        if let v = parsed.subject  { conditions.append("subject LIKE ?");       bindings.append("%\(v)%") }
        if let v = parsed.id       { conditions.append("id = ?");               bindings.append(v) }
        if parsed.hasAttachment    { conditions.append("attachment_count > 0") }
        if let v = parsed.freeText {
            conditions.append("(subject LIKE ? OR sender LIKE ? OR to_recipients LIKE ? OR body_text LIKE ?)")
            let t = "%\(v)%"; bindings += [t, t, t, t]
        }
        if let d = lastFromDate {
            conditions.append("COALESCE(delivery_time, submit_time) >= ?")
            bindings.append(Self.sqlDateFormatter.string(from: d))
        }
        if let d = lastToDate {
            // Use start of the following day so the entire toDate is included.
            let next = Calendar.current.date(byAdding: .day, value: 1, to: d) ?? d
            conditions.append("COALESCE(delivery_time, submit_time) < ?")
            bindings.append(Self.sqlDateFormatter.string(from: next))
        }

        let where_ = conditions.isEmpty ? "" : "WHERE " + conditions.joined(separator: " AND ")
        let orderCol: String
        switch lastSort {
        case .deliveryTime: orderCol = "delivery_time"
        case .submitTime:   orderCol = "submit_time"
        case .sender:       orderCol = "sender"
        case .recipient:    orderCol = "to_recipients"
        case .subject:      orderCol = "subject"
        }
        let dir = lastAscending ? "ASC" : "DESC"
        let sql = """
            SELECT id, folder_id, subject, sender, to_recipients, cc_recipients,
                   submit_time, delivery_time, body_text, body_html, attachment_count
            FROM messages \(where_)
            ORDER BY \(orderCol) \(dir)
            LIMIT \(pageSize) OFFSET \(offset)
            """

        log("execute sql (offset=\(offset)): \(sql.trimmingCharacters(in: .whitespacesAndNewlines))")
        var stmt: OpaquePointer?
        guard sqlite3_prepare_v2(db, sql, -1, &stmt, nil) == SQLITE_OK else {
            log("execute prepare failed: \(sqliteError())")
            return
        }
        defer { sqlite3_finalize(stmt) }

        let SQLITE_TRANSIENT = unsafeBitCast(-1, to: sqlite3_destructor_type.self)
        for (i, val) in bindings.enumerated() {
            let idx = Int32(i + 1)
            if let s = val as? String {
                sqlite3_bind_text(stmt, idx, (s as NSString).utf8String, -1, SQLITE_TRANSIENT)
            } else if let n = val as? Int64 {
                sqlite3_bind_int64(stmt, idx, n)
            }
        }

        let iso = ISO8601DateFormatter()
        var result: [Message] = []
        while sqlite3_step(stmt) == SQLITE_ROW {
            func str(_ col: Int32) -> String {
                guard let p = sqlite3_column_text(stmt, col) else { return "" }
                return String(cString: p)
            }
            func optStr(_ col: Int32) -> String? {
                guard sqlite3_column_type(stmt, col) != SQLITE_NULL,
                      let p = sqlite3_column_text(stmt, col) else { return nil }
                return String(cString: p)
            }
            result.append(Message(
                id:              sqlite3_column_int64(stmt, 0),
                folderId:        sqlite3_column_int64(stmt, 1),
                subject:         str(2).isEmpty ? "(no subject)" : str(2),
                sender:          str(3).isEmpty ? "(unknown)" : str(3),
                toRecipients:    str(4),
                ccRecipients:    str(5),
                submitTime:      optStr(6).flatMap { iso.date(from: $0) },
                deliveryTime:    optStr(7).flatMap { iso.date(from: $0) },
                bodyText:        optStr(8),
                bodyHtml:        optStr(9),
                attachmentCount: Int(sqlite3_column_int(stmt, 10))
            ))
        }
        let more = result.count == pageSize
        log("execute: got \(result.count) message(s), hasMore=\(more)")
        DispatchQueue.main.async {
            if appending {
                self.messages.append(contentsOf: result)
            } else {
                self.messages = result
            }
            self.hasMore = more
        }
    }

    func loadAttachments(messageId: Int64) -> [Attachment] {
        guard let db else { log("loadAttachments: no open db"); return [] }
        let sql = "SELECT id, message_id, filename, content_type, size, data FROM attachments WHERE message_id = ?"
        var stmt: OpaquePointer?
        guard sqlite3_prepare_v2(db, sql, -1, &stmt, nil) == SQLITE_OK else {
            log("loadAttachments prepare failed: \(sqliteError())")
            return []
        }
        defer { sqlite3_finalize(stmt) }
        sqlite3_bind_int64(stmt, 1, messageId)
        var result: [Attachment] = []
        while sqlite3_step(stmt) == SQLITE_ROW {
            func str(_ col: Int32) -> String {
                guard let p = sqlite3_column_text(stmt, col) else { return "" }
                return String(cString: p)
            }
            let dataPtr = sqlite3_column_blob(stmt, 5)
            let dataLen = sqlite3_column_bytes(stmt, 5)
            let data: Data? = dataPtr.map { Data(bytes: $0, count: Int(dataLen)) }
            result.append(Attachment(
                id:          sqlite3_column_int64(stmt, 0),
                messageId:   sqlite3_column_int64(stmt, 1),
                filename:    str(2).isEmpty ? "attachment" : str(2),
                contentType: str(3),
                size:        Int(sqlite3_column_int(stmt, 4)),
                data:        data
            ))
        }
        return result
    }
}

// MARK: - Sort

enum SortField: String, CaseIterable, Identifiable {
    case deliveryTime = "Received"
    case submitTime   = "Sent"
    case sender       = "Sender"
    case recipient    = "Recipient"
    case subject      = "Subject"
    var id: Self { self }
}

// MARK: - App Entry Point

@main
struct PSTBrowserApp: App {
    var body: some Scene {
        WindowGroup {
            ContentView()
                .frame(minWidth: 900, minHeight: 600)
        }
        .commands {
            CommandGroup(replacing: .newItem) {}
        }
    }
}

// MARK: - Content View

struct ContentView: View {
    @StateObject private var db = Database()
    @State private var selectedFolder: Folder?
    @State private var selectedMessage: Message?
    @State private var search = ""
    @State private var sort: SortField = .deliveryTime
    @State private var ascending = false
    @State private var fromDate: Date? = nil
    @State private var toDate: Date? = nil
    @State private var showOpenPanel = false
    @State private var sidebarWidth: CGFloat = 200

    var body: some View {
        Group {
            if db.isOpen {
                mainLayout
            } else {
                welcomeView
            }
        }
        .frame(minWidth: 900, minHeight: 600)
        .onAppear { showOpenPanel = !db.isOpen }
    }

    // MARK: Welcome

    private var welcomeView: some View {
        VStack(spacing: 20) {
            Image(systemName: "envelope.open")
                .font(.system(size: 64))
                .foregroundColor(.secondary)
            Text("PST Browser")
                .font(.largeTitle.bold())
            Text("Open a pstexplorer SQLite export to get started.")
                .foregroundColor(.secondary)
            Button("Open Database…") { openFile() }
                .buttonStyle(.borderedProminent)
                .controlSize(.large)
            if let err = db.error {
                Text(err).foregroundColor(.red).font(.caption)
            }
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
    }

    // MARK: Main layout

    private var mainLayout: some View {
        NavigationSplitView {
            sidebarView
                .navigationSplitViewColumnWidth(min: 150, ideal: 200, max: 300)
        } content: {
            messageListView
                .navigationSplitViewColumnWidth(min: 260, ideal: 340, max: 500)
        } detail: {
            if let msg = selectedMessage {
                MessageDetailView(message: msg, db: db)
            } else {
                Text("Select a message")
                    .foregroundColor(.secondary)
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
            }
        }
        .toolbar {
            ToolbarItem(placement: .navigation) {
                Button(action: openFile) {
                    Label("Open…", systemImage: "folder")
                }
                .help("Open a different database")
            }
            ToolbarItem(placement: .primaryAction) {
                HStack {
                    Picker("Sort by", selection: $sort) {
                        ForEach(SortField.allCases) { f in
                            Text(f.rawValue).tag(f)
                        }
                    }
                    .pickerStyle(.menu)
                    .frame(width: 130)
                    Button {
                        ascending.toggle()
                        reload()
                    } label: {
                        Image(systemName: ascending ? "arrow.up" : "arrow.down")
                    }
                    .help(ascending ? "Ascending" : "Descending")
                }
            }
            ToolbarItem(placement: .primaryAction) {
                TextField("Search… from: to: subject:", text: $search)
                    .frame(width: 180)
                    .textFieldStyle(.roundedBorder)
                    .onSubmit { reload() }
            }
        }
        .onChange(of: sort)           { reload() }
        .onChange(of: selectedFolder) { reload() }
        .onChange(of: fromDate)       { reload() }
        .onChange(of: toDate)         { reload() }
    }

    // MARK: Sidebar

    private var sidebarView: some View {
        List(selection: $selectedFolder) {
            Section("Folders") {
                Button {
                    selectedFolder = nil
                } label: {
                    Label("All Mail", systemImage: "tray.full")
                        .foregroundColor(selectedFolder == nil ? .accentColor : .primary)
                }
                .buttonStyle(.plain)

                ForEach(db.folders) { folder in
                    let depth = folder.path.components(separatedBy: "/").count - 2
                    Label(folder.name, systemImage: folderIcon(for: folder.name))
                        .padding(.leading, CGFloat(depth) * 12)
                        .tag(folder)
                }
            }
        }
        .listStyle(.sidebar)
        .navigationTitle("PST Browser")
    }

    private func folderIcon(for name: String) -> String {
        let lower = name.lowercased()
        if lower.contains("inbox")    { return "tray" }
        if lower.contains("sent")     { return "paperplane" }
        if lower.contains("draft")    { return "doc" }
        if lower.contains("trash") || lower.contains("deleted") { return "trash" }
        if lower.contains("junk") || lower.contains("spam")     { return "xmark.bin" }
        if lower.contains("archive")  { return "archivebox" }
        return "folder"
    }

    // MARK: Message list

    private var dateFilterBar: some View {
        HStack(spacing: 6) {
            Image(systemName: "calendar")
                .foregroundColor(.secondary)
                .font(.caption)
            OptionalDatePicker(label: "From", date: $fromDate)
            Text("→")
                .foregroundColor(.secondary)
                .font(.caption)
            OptionalDatePicker(label: "To", date: $toDate)
            if fromDate != nil || toDate != nil {
                Button { fromDate = nil; toDate = nil } label: {
                    Image(systemName: "xmark.circle.fill")
                        .foregroundColor(.secondary)
                }
                .buttonStyle(.plain)
            }
            Spacer()
        }
        .padding(.horizontal, 10)
        .padding(.vertical, 5)
        .background(Color(.windowBackgroundColor))
        .overlay(alignment: .bottom) { Divider() }
    }

    private var messageListView: some View {
        List(selection: $selectedMessage) {
            ForEach(db.messages) { msg in
                MessageRowView(message: msg)
                    .tag(msg)
                    .listRowInsets(EdgeInsets(top: 6, leading: 10, bottom: 6, trailing: 10))
                    .onAppear {
                        if db.hasMore && msg == db.messages.last {
                            db.loadNextPage()
                        }
                    }
            }
            if db.hasMore {
                ProgressView()
                    .frame(maxWidth: .infinity)
                    .listRowSeparator(.hidden)
            }
        }
        .listStyle(.plain)
        .navigationTitle(selectedFolder?.name ?? "All Mail")
        .safeAreaInset(edge: .top, spacing: 0) { dateFilterBar }
        .overlay {
            if db.messages.isEmpty {
                Text("No messages")
                    .foregroundColor(.secondary)
            }
        }
    }

    // MARK: Helpers

    private func reload() {
        db.loadMessages(folderId: selectedFolder?.id, search: search, sort: sort,
                        ascending: ascending, fromDate: fromDate, toDate: toDate)
        selectedMessage = nil
    }

    private func openFile() {
        let panel = NSOpenPanel()
        panel.title = "Open PST Export Database"
        panel.allowedContentTypes = [.init(filenameExtension: "db")!, .init(filenameExtension: "sqlite")!, .init(filenameExtension: "sqlite3")!]
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        if panel.runModal() == .OK, let url = panel.url {
            db.open(url: url)
            selectedFolder = nil
            selectedMessage = nil
            fromDate = nil
            toDate   = nil
        }
    }
}

// MARK: - Optional Date Picker

struct OptionalDatePicker: View {
    let label: String
    @Binding var date: Date?

    var body: some View {
        if let d = date {
            HStack(spacing: 2) {
                DatePicker(
                    "",
                    selection: Binding(get: { d }, set: { date = $0 }),
                    displayedComponents: .date
                )
                .datePickerStyle(.compact)
                .labelsHidden()
                .fixedSize()

                Button { date = nil } label: {
                    Image(systemName: "xmark.circle.fill")
                        .font(.caption)
                        .foregroundColor(.secondary)
                }
                .buttonStyle(.plain)
            }
        } else {
            Button(label) { date = Date() }
                .buttonStyle(.plain)
                .foregroundColor(.accentColor)
                .font(.subheadline)
        }
    }
}

// MARK: - Message Row

struct MessageRowView: View {
    let message: Message
    private static let dateFormatter: DateFormatter = {
        let f = DateFormatter()
        f.doesRelativeDateFormatting = true
        f.dateStyle = .short
        f.timeStyle = .short
        return f
    }()

    var body: some View {
        VStack(alignment: .leading, spacing: 3) {
            HStack {
                Text(message.sender)
                    .font(.headline)
                    .lineLimit(1)
                Spacer()
                Text(formattedDate)
                    .font(.caption)
                    .foregroundColor(.secondary)
            }
            Text(message.subject)
                .font(.subheadline)
                .lineLimit(1)
                .foregroundColor(.primary)
            HStack(spacing: 4) {
                Text(message.bodyText?.trimmingCharacters(in: .whitespacesAndNewlines).prefix(80).replacingOccurrences(of: "\n", with: " ") ?? "")
                    .font(.caption)
                    .foregroundColor(.secondary)
                Spacer()
                if message.attachmentCount > 0 {
                    Image(systemName: "paperclip")
                        .font(.caption)
                        .foregroundColor(.secondary)
                }
            }
        }
        .padding(.vertical, 2)
    }

    private var formattedDate: String {
        let date = message.deliveryTime ?? message.submitTime
        return date.map { Self.dateFormatter.string(from: $0) } ?? ""
    }
}

// MARK: - Message Detail

struct MessageDetailView: View {
    let message: Message
    let db: Database
    @State private var attachments: [Attachment] = []
    @State private var showHtml: Bool = true

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 0) {
                headerSection
                Divider()
                bodySection
                if !attachments.isEmpty {
                    Divider()
                    attachmentSection
                }
            }
            .padding()
        }
        .onAppear {
            if message.attachmentCount > 0 {
                attachments = db.loadAttachments(messageId: message.id)
            }
        }
        .onChange(of: message) {
            attachments = message.attachmentCount > 0 ? db.loadAttachments(messageId: message.id) : []
            showHtml = true
        }
    }

    // MARK: Header

    private var headerSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(message.subject)
                .font(.title2.bold())
                .textSelection(.enabled)

            Divider()

            HeaderRow(label: "From",    value: message.sender)
            HeaderRow(label: "To",      value: message.toRecipients)
            if !message.ccRecipients.isEmpty {
                HeaderRow(label: "CC",  value: message.ccRecipients)
            }
            if let date = message.submitTime {
                HeaderRow(label: "Sent", value: fullDateFormatter.string(from: date))
            }
            if let date = message.deliveryTime {
                HeaderRow(label: "Received", value: fullDateFormatter.string(from: date))
            }
            HeaderRow(label: "ID", value: String(message.id))
        }
        .padding(.bottom, 12)
    }

    // MARK: Body

    private var bodySection: some View {
        let hasHtml = !(message.bodyHtml ?? "").isEmpty
        let hasText = !(message.bodyText ?? "").isEmpty
        return VStack(alignment: .trailing, spacing: 8) {
            if hasHtml && hasText {
                Picker("View", selection: $showHtml) {
                    Text("HTML").tag(true)
                    Text("Plain text").tag(false)
                }
                .pickerStyle(.segmented)
                .fixedSize()
                .padding(.top, 8)
            }
            if hasHtml && showHtml {
                WebView(html: message.bodyHtml!)
                    .frame(minHeight: 300)
            } else if hasText {
                Text(message.bodyText!)
                    .font(.body)
                    .textSelection(.enabled)
                    .frame(maxWidth: .infinity, alignment: .leading)
            } else {
                Text("(no content)")
                    .foregroundColor(.secondary)
                    .frame(maxWidth: .infinity, alignment: .leading)
            }
        }
        .padding(.top, 8)
    }

    // MARK: Attachments

    private var attachmentSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Attachments (\(attachments.count))")
                .font(.headline)
                .padding(.top, 12)

            FlowLayout(spacing: 8) {
                ForEach(attachments) { att in
                    AttachmentChip(attachment: att)
                }
            }
        }
    }

    private var fullDateFormatter: DateFormatter {
        let f = DateFormatter()
        f.dateStyle = .long
        f.timeStyle = .medium
        return f
    }
}

// MARK: - Header Row

struct HeaderRow: View {
    let label: String
    let value: String

    var body: some View {
        HStack(alignment: .top, spacing: 0) {
            Text(label + ":")
                .font(.subheadline.bold())
                .foregroundColor(.secondary)
                .frame(width: 68, alignment: .trailing)
            Text(value)
                .font(.subheadline)
                .textSelection(.enabled)
                .padding(.leading, 8)
                .frame(maxWidth: .infinity, alignment: .leading)
        }
    }
}

// MARK: - Attachment Chip

struct AttachmentChip: View {
    let attachment: Attachment
    @State private var isHovered = false

    var body: some View {
        Button(action: saveAttachment) {
            HStack(spacing: 6) {
                Image(systemName: iconForMimeType(attachment.contentType))
                    .font(.body)
                VStack(alignment: .leading, spacing: 1) {
                    Text(attachment.filename)
                        .font(.caption.bold())
                        .lineLimit(1)
                    Text(formatSize(attachment.size))
                        .font(.caption2)
                        .foregroundColor(.secondary)
                }
            }
            .padding(.horizontal, 10)
            .padding(.vertical, 6)
            .background(isHovered ? Color.accentColor.opacity(0.15) : Color(.controlBackgroundColor))
            .cornerRadius(8)
            .overlay(
                RoundedRectangle(cornerRadius: 8)
                    .stroke(Color(.separatorColor), lineWidth: 1)
            )
        }
        .buttonStyle(.plain)
        .onHover { isHovered = $0 }
        .help("Save \(attachment.filename)")
    }

    private func saveAttachment() {
        guard let data = attachment.data else { return }
        let panel = NSSavePanel()
        panel.nameFieldStringValue = attachment.filename
        if panel.runModal() == .OK, let url = panel.url {
            try? data.write(to: url)
        }
    }

    private func iconForMimeType(_ mime: String) -> String {
        if mime.hasPrefix("image/")       { return "photo" }
        if mime.hasPrefix("video/")       { return "video" }
        if mime.hasPrefix("audio/")       { return "music.note" }
        if mime.contains("pdf")           { return "doc.richtext" }
        if mime.contains("zip") || mime.contains("compressed") { return "archivebox" }
        if mime.contains("word") || mime.contains("msword")    { return "doc.text" }
        if mime.contains("spreadsheet") || mime.contains("excel") { return "tablecells" }
        return "paperclip"
    }

    private func formatSize(_ bytes: Int) -> String {
        let kb = Double(bytes) / 1024
        if kb < 1024 { return String(format: "%.0f KB", kb) }
        return String(format: "%.1f MB", kb / 1024)
    }
}

// MARK: - WebView (WKWebView wrapper)

struct WebView: NSViewRepresentable {
    let html: String

    func makeNSView(context: Context) -> WKWebView {
        let config = WKWebViewConfiguration()
        config.preferences.setValue(true, forKey: "allowFileAccessFromFileURLs")
        let webView = WKWebView(frame: .zero, configuration: config)
        webView.setValue(false, forKey: "drawsBackground")
        return webView
    }

    func updateNSView(_ webView: WKWebView, context: Context) {
        // Wrap in a minimal HTML shell that respects system appearance
        let wrapped = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light dark">
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            font-size: 14px;
            line-height: 1.5;
            margin: 0;
            padding: 12px 0;
            color: -apple-system-label;
            background: transparent;
            word-wrap: break-word;
          }
          a { color: -apple-system-blue; }
          img { max-width: 100%; height: auto; }
          pre, code { white-space: pre-wrap; font-size: 12px; }
        </style>
        </head>
        <body>\(html)</body>
        </html>
        """
        webView.loadHTMLString(wrapped, baseURL: nil)
    }
}

// MARK: - Flow Layout (wrapping HStack for attachment chips)

struct FlowLayout: Layout {
    var spacing: CGFloat = 8

    func sizeThatFits(proposal: ProposedViewSize, subviews: Subviews, cache: inout ()) -> CGSize {
        let width = proposal.width ?? .infinity
        var x: CGFloat = 0, y: CGFloat = 0, rowHeight: CGFloat = 0, maxX: CGFloat = 0
        for view in subviews {
            let size = view.sizeThatFits(.unspecified)
            if x + size.width > width && x > 0 {
                y += rowHeight + spacing; x = 0; rowHeight = 0
            }
            rowHeight = max(rowHeight, size.height)
            x += size.width + spacing
            maxX = max(maxX, x)
        }
        return CGSize(width: maxX, height: y + rowHeight)
    }

    func placeSubviews(in bounds: CGRect, proposal: ProposedViewSize, subviews: Subviews, cache: inout ()) {
        var x = bounds.minX, y = bounds.minY, rowHeight: CGFloat = 0
        for view in subviews {
            let size = view.sizeThatFits(.unspecified)
            if x + size.width > bounds.maxX && x > bounds.minX {
                y += rowHeight + spacing; x = bounds.minX; rowHeight = 0
            }
            view.place(at: CGPoint(x: x, y: y), proposal: ProposedViewSize(size))
            rowHeight = max(rowHeight, size.height)
            x += size.width + spacing
        }
    }
}
