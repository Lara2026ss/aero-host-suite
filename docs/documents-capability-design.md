# FLUXER XZ — Documents Capability Architectural Design Document

**Version**: 1.0.0  
**Domain**: `documents`  
**Target Systems**: Microsoft Word, Microsoft PowerPoint, Microsoft Excel (Desktop Office 365 on Windows 11)  
**Author**: Executive Engineer  
**Status**: APPROVED FOR IMPLEMENTATION  

---

## 1. Executive Summary & Design Goals

The `documents` capability provides FLUXER XZ with first-class, real-time control over Microsoft Office desktop applications on Windows 11. It allows agents to inspect, create, read, and edit documents in **real time while open on the user's screen**, as well as operate in headless mode when applications are closed.

### Core Tenets
1. **LIVE Visibility First**: When the user has Word, PowerPoint, or Excel open, edits attach directly to the running instance via Windows COM automation (`GetActiveObject`), rendering changes immediately without reopening files.
2. **File Lock & Integrity Protection**: Direct byte manipulation on disk while an Office application holds a lock leads to file corruption or sharing violations (`EBUSY`). The engine routes all modifications for open files through the live COM object.
3. **No Native Add-on Overhead**: Achieved entirely using Windows-native COM automation executed via child PowerShell processes communicating over a secure JSON IPC channel. Zero new npm binary dependencies.
4. **Injection-Proof IPC**: User content, file paths, and replacement strings are never interpolated directly into PowerShell command lines; they are transferred via temporary structured JSON IPC files with strict cleanup.
5. **Token Efficiency**: Default outputs return structural summaries (heading outlines, paragraph counts, slide counts) and before/after diff snippets rather than dumping 100-page raw text documents into the LLM context.

---

## 2. Architecture & Operational Modes

```
                                [Client Tool Invocation]
                                           │
                                           ▼
                             [tools/documents.mjs]
                         (Validation, Permissions & Guard)
                                           │
                        Is document or target app active?
                                   /       \
                                  /         \
                             YES /           \ NO
                                /             \
                               ▼               ▼
                      [LIVE COM MODE]      [HEADLESS COM MODE / OOXML]
                     Attach to active      Launch hidden COM instance
                     process in memory     or load via docx/exceljs
                               \               /
                                \             /
                                 ▼           ▼
                      [tools/documents/office-com.ps1]
                           (Safe JSON IPC Channel)
                                           │
                                           ▼
                     [Active Desktop App / Target File System]
```

### 2.1 Mode Resolution Matrix
- **LIVE Mode**: Target document is already open in Word, PowerPoint, or Excel. The engine attaches via `[System.Runtime.InteropServices.Marshal]::GetActiveObject(...)` or through application document collections (`$word.Documents`, `$ppt.Presentations`, `$excel.Workbooks`).
- **HEADLESS COM Mode**: Application is not running or document is closed. Launches a hidden background COM instance (`New-Object -ComObject Word.Application`), loads the file, performs mutations, saves, and cleanly releases COM memory.
- **STATIC FALLBACK Mode**: When lightweight file creation or PDF reading is requested and Office COM is unneeded, fall back to already-present dependencies (`docx`, `exceljs`, `pdf-lib`, `mammoth`).

---

## 3. Action Catalog & Implementation Status

> **Inventory Accounting**:
> - **Total Inventoried Actions**: 593 actions across 21 registered domains (including 23 alias actions from the canonical `docs` alias).
> - **Unique Actions**: 570 actions across unique capability modules (593 minus 23 duplicate `docs` alias actions).

### 3.1 Discovery Actions
1. `status` **[IMPLEMENTED]**: Inspects whether Microsoft Word, PowerPoint, and Excel are installed, running, and accessible via COM.
   - Parameters: `{ app?: "word" | "powerpoint" | "excel" | "all" }`
   - Returns: Status of each application (installed, running, processId, version).
2. `list_open` **[IMPLEMENTED]**: Enumerates all documents currently open in memory across running Office instances.
   - Parameters: `{ app?: "word" | "powerpoint" | "excel" | "all" }`
   - Returns: Array of `{ app, title, path, saved, active, pages/slides/sheets }`.
3. `get_active` **[IMPLEMENTED]**: Retrieves the active document in focus.
   - Parameters: `{ app: "word" | "powerpoint" | "excel" }`
   - Returns: Metadata of active document.

### 3.2 Lifecycle Actions
4. `create` **[IMPLEMENTED]**: Creates a new blank or template-based document/presentation/workbook.
   - Parameters: `{ app: "word" | "powerpoint" | "excel", path: string, template?: string, visible?: boolean, overwrite?: boolean }`
   - Returns: `{ ok: true, path, app, opened: boolean }`.
5. `open` **[IMPLEMENTED]**: Opens an existing file in desktop Office.
   - Parameters: `{ path: string, visible?: boolean, readOnly?: boolean }`
   - Returns: `{ ok: true, path, app, readOnly }`.
6. `save` **[IMPLEMENTED]**: Saves an open document in-place.
   - Parameters: `{ path?: string, app?: string }`
   - Returns: `{ ok: true, path, saved: true }`.
7. `save_as` **[IMPLEMENTED]**: Saves document under a new path or format.
   - Parameters: `{ path: string, targetPath: string, format?: "docx"|"pptx"|"xlsx"|"pdf"|"html" }`
   - Returns: `{ ok: true, sourcePath, targetPath }`.
8. `close` **[IMPLEMENTED]**: Closes a document safely.
   - Parameters: `{ path: string, saveChanges?: boolean, closeApp?: boolean }`
   - Returns: `{ ok: true, path, closed: true }`.
9. `export_pdf` **[IMPLEMENTED]**: Direct one-click PDF export using native Office print/export engines.
   - Parameters: `{ path: string, outputPath?: string }`
   - Returns: `{ ok: true, pdfPath, sizeBytes }`.

### 3.3 Word Specific Actions
10. `read` **[IMPLEMENTED]**: Structured read of a Word document (headings, paragraph summary, table counts, comments).
    - Parameters: `{ path: string, maxChars?: number, range?: string, includeComments?: boolean }`
11. `insert_text` **[IMPLEMENTED]**: Inserts text at beginning, end, or relative position.
    - Parameters: `{ path: string, text: string, position?: "end"|"start"|"selection", style?: string, bold?: boolean, italic?: boolean, fontSize?: number }`
12. `replace_text` **[IMPLEMENTED]**: Search and replace strings with optional matchCase and wholeWord.
    - Parameters: `{ path: string, search: string, replace: string, matchCase?: boolean, matchWholeWord?: boolean, replaceAll?: boolean }`
13. `append_section` **[IMPLEMENTED]**: Adds a formatted section with a heading and body paragraphs.
    - Parameters: `{ path: string, heading: string, body: string, headingLevel?: 1|2|3 }`
14. `insert_table` **[IMPLEMENTED]**: Inserts a structured data table.
    - Parameters: `{ path: string, rows: number, cols: number, data?: Array<Array<string>>, headerRow?: boolean }`
15. `insert_image` **[IMPLEMENTED]**: Inserts an image into the document at inline shape position.
    - Parameters: `{ path: string, imagePath: string, width?: number, height?: number }`
16. `add_comment` **[PLANIFICADO / ROADMAP]**: Adds a comment to the current selection or specified text range.
    - Parameters: `{ path: string, text: string, comment: string }`
17. `set_header_footer` **[PLANIFICADO / ROADMAP]**: Configures header and footer text.
    - Parameters: `{ path: string, header?: string, footer?: string }`
18. `track_changes` **[PLANIFICADO / ROADMAP]**: Enables or disables change tracking.
    - Parameters: `{ path: string, enabled: boolean }`
19. `set_style` **[PLANIFICADO / ROADMAP]**: Applies custom or predefined typography and character styles.
    - Parameters: `{ path: string, range?: string, styleName: string }`

### 3.4 PowerPoint Specific Actions
20. `read_slides` **[IMPLEMENTED]**: Extracts slide list, slide titles, shapes, and presenter notes.
    - Parameters: `{ path: string, slideIndex?: number }`
21. `add_slide` **[IMPLEMENTED]**: Adds a new slide with specified layout (`title`, `title_content`, `blank`, `section`) and text.
    - Parameters: `{ path: string, layout?: "title"|"title_content"|"blank"|"section", title?: string, body?: string }`
22. `set_slide_text` **[IMPLEMENTED]**: Sets or replaces text within a specific slide shape.
    - Parameters: `{ path: string, slideIndex: number, shapeIndex?: number, text: string }`
    - Returns: `{ ok: true, slideIndex, shapeIndex, updatedText }`
23. `insert_shape` **[IMPLEMENTED]**: Inserts shapes (rectangles, callouts, text boxes) onto a slide.
    - Parameters: `{ path: string, slideIndex: number, shapeType: string, left: number, top: number, width: number, height: number, text?: string }`
24. `reorder_slides` **[PLANIFICADO / ROADMAP]**: Moves slide from one index to another.
    - Parameters: `{ path: string, fromIndex: number, toIndex: number }`
25. `delete_slide` **[IMPLEMENTED]**: Deletes a slide by index.
    - Parameters: `{ path: string, slideIndex: number }`
26. `apply_theme` **[PLANIFICADO / ROADMAP]**: Applies a design template or theme to a presentation.
    - Parameters: `{ path: string, themePath: string }`
27. `set_notes` **[IMPLEMENTED]**: Sets presenter notes on a slide.
    - Parameters: `{ path: string, slideIndex: number, notes: string }`

### 3.5 Excel Specific Actions
28. `read_range` **[IMPLEMENTED]**: Reads values and formulas from a designated worksheet range.
    - Parameters: `{ path: string, sheet?: string|number, range?: string }`
29. `write_range` **[IMPLEMENTED]**: Writes values or formulas to a worksheet range.
    - Parameters: `{ path: string, sheet?: string|number, startCell: string, data: Array<Array<any>> }`
30. `add_sheet` **[PLANIFICADO / ROADMAP]**: Adds a new worksheet with an optional name.
    - Parameters: `{ path: string, name?: string }`
31. `set_formula` **[PLANIFICADO / ROADMAP]**: Sets a formula in a designated cell.
    - Parameters: `{ path: string, sheet?: string|number, cell: string, formula: string }`

### 3.6 Batch Operations
32. `batch_create` **[IMPLEMENTED]**: Generates multiple documents from an array of specifications in a single call (with `overwrite: false` by default).
    - Parameters: `{ documents: Array<{ app, path, title, content, overwrite?: boolean }> }`
33. `batch_edit` **[IMPLEMENTED]**: Applies an ordered series of operations to a document, stopping and reporting on failure.
    - Parameters: `{ path: string, app: string, operations: Array<{ action: string, args: object }> }`

---

## 4. Security, IPC & Resilience Architecture

### 4.1 Injection-Proof Child Process IPC
```javascript
// High-level pattern in tools/documents.mjs
const tempIpcPath = path.join(runtime.dirs.storage, `doc_ipc_${crypto.randomUUID()}.json`);
try {
  await fs.writeFile(tempIpcPath, JSON.stringify(payload, null, 2), "utf8");
  const { stdout, stderr } = await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy", "Bypass",
    "-File", psScriptPath,
    "-PayloadFile", tempIpcPath
  ], { timeout: 45000, windowsHide: true });
  return JSON.parse(stdout);
} finally {
  await fs.unlink(tempIpcPath).catch(() => {});
}
```

### 4.2 COM Error Recovery & Retries
- **RPC Server Busy (`0x8001010A`)**: When Office is computing or waiting for UI focus, PowerShell catches the COM exception, pauses for 500ms, and retries once.
- **Modal Dialog Lock**: If a user has a modal "Save As" or "Font" dialog open, the COM call returns code `OFFICE_DIALOG_OPEN` with a clear advisory to the agent rather than hanging indefinitely.
- **Protected View / Read-Only**: Documents opened from internet zones are flagged with code `OFFICE_PROTECTED_VIEW`.

### 4.3 Permissions Mapping
- `read`, `status`, `list_open`, `get_active`, `read_slides`, `read_range`: **standard (user level)**.
- `create`, `open`, `save`, `save_as`, `insert_text`, `replace_text`, `append_section`, `insert_table`, `insert_image`, `add_comment`, `set_header_footer`, `track_changes`, `add_slide`, `set_slide_text`, `insert_shape`, `reorder_slides`, `delete_slide`, `apply_theme`, `set_notes`, `write_range`, `add_sheet`, `set_formula`, `batch_create`, `batch_edit`: **standard (files write policy)**.
- `close` (with discard / `saveChanges: false` on unsaved edits), destructive overwrites without backup: **advanced (power_user/operator policy)**.

---

## 5. Doctor & Test Integration Strategy

1. **Discovery**: Add `documents` to `CAPABILITY_NAMES` and `CAPABILITY_DEFINITIONS` in `core/capability-registry.mjs`.
2. **Registry Loader**: Import `createDocumentsDomain` in `core/registry.mjs`.
3. **Doctor Verification**: Update discovery expectations so `doctor.mjs` discovers 20 domains and the increased action count without regressions.
4. **Integration Test Suite**: Add comprehensive test cases to `tests/test_fluxer_suite.mjs` validating real document creation, inspection, and cleanup.
