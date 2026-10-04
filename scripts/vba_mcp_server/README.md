# VBA MCP Server

**Let AI read and write your Excel VBA code directly — no manual export/import needed.**

English | [中文](README.zh-CN.md) | [日本語](README.ja.md)

> An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude (and other AI assistants) the ability to directly access and modify VBA code inside Excel workbooks (`.xlsm`, `.xlsb`, `.xls`, `.xla`, `.xlam`).

---

## Why This Exists

If you've ever worked with Excel VBA and AI coding assistants, you know the pain:

1. Open VBA Editor, manually copy code
2. Paste into chat, ask AI for help
3. Copy AI's response back
4. Paste into VBA Editor, test, repeat...

**VBA MCP Server eliminates this entirely.** Claude can now read, write, create, delete, and backup VBA modules directly — just like it works with regular source code files.

---

## Features

- **Direct VBA Access** — Read and write VBA code in-place, no export/import cycle
- **Smart Workbook Handling** — Detects already-open workbooks, avoids conflicts
- **Auto-Save** — Workbooks are saved automatically after writes
- **Full Module Management** — Create, delete, read, write, list, and backup modules
- **Workbook Discovery** — Find all `.xlsm` files in a directory tree
- **Backup with Manifest** — Export all modules to text files with JSON metadata
- **Document Module Protection** — Prevents accidental deletion of Sheet/ThisWorkbook modules
- **Lightweight** — Single Python file (~300 lines), two dependencies

---

## Prerequisites

| Requirement | Details |
|------------|---------|
| **OS** | Windows (uses COM/Win32 API) |
| **Excel** | Microsoft Excel (local installation) |
| **Python** | 3.10 or later |
| **Trust Setting** | See [Excel Configuration](#excel-configuration) below |

---

## Installation

### Option 1: pip install (Recommended)

```bash
pip install vba-mcp-server
```

Then add to your MCP config:

```json
{
  "mcpServers": {
    "vba-mcp-server": {
      "command": "vba-mcp-server",
      "args": []
    }
  }
}
```

### Option 2: Claude Code (Manual)

Add to your Claude Code MCP settings (`~/.claude.json` or project `.mcp.json`):

```json
{
  "mcpServers": {
    "vba-mcp-server": {
      "command": "python",
      "args": ["C:/path/to/vba-mcp-server/server.py"],
      "env": {}
    }
  }
}
```

### Option 3: Claude Desktop

Add to your Claude Desktop config (`%APPDATA%/Claude/claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "vba-mcp-server": {
      "command": "python",
      "args": ["C:/path/to/vba-mcp-server/server.py"]
    }
  }
}
```

### Option 4: From Source

```bash
# Clone the repository
git clone https://github.com/xiongchenghou/vba-mcp-server.git
cd vba-mcp-server

# Create virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## Excel Configuration

**Required one-time setup** — Excel must allow programmatic access to VBA:

1. Open Excel
2. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
3. Click **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click **OK**

> Without this setting, all tools will return an error with a helpful hint message.

---

## Available Tools

### `vba_list_modules`

List all VBA modules in a workbook with metadata.

```
→ vba_list_modules("C:/Projects/MyWorkbook.xlsm")

[
  {"name": "Module1", "type": "StandardModule", "line_count": 342},
  {"name": "Sheet1",  "type": "Document",       "line_count": 15},
  {"name": "MyClass", "type": "ClassModule",     "line_count": 89}
]
```

### `vba_read_module`

Read the source code of a specific module.

```
→ vba_read_module("C:/Projects/MyWorkbook.xlsm", "Module1")

' === Module: Module1 (StandardModule) ===
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub
```

### `vba_read_all`

Read ALL modules at once — perfect for full codebase review.

```
→ vba_read_all("C:/Projects/MyWorkbook.xlsm")

============================================================
' Module: Module1
' Type: StandardModule
' Lines: 342
============================================================
[full source code...]

============================================================
' Module: Sheet1
' Type: Document
' Lines: 15
============================================================
[full source code...]
```

### `vba_write_module`

Write or replace a module's code. Auto-saves the workbook. Creates the module if it doesn't exist.

```
→ vba_write_module("C:/Projects/MyWorkbook.xlsm", "Module1", "Sub Test()\n    MsgBox \"Updated!\"\nEnd Sub")

{"success": true, "module": "Module1", "lines_written": 3, "message": "Module 'Module1' updated (3 lines)"}
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `file_path` | string | required | Path to the workbook |
| `module_name` | string | required | Name of the module to write |
| `code` | string | required | VBA source code |
| `create_if_missing` | bool | `true` | Create module if it doesn't exist |

### `vba_create_module`

Create a new VBA module (standard or class).

```
→ vba_create_module("C:/Projects/MyWorkbook.xlsm", "MyNewModule", "Option Explicit", "standard")

{"success": true, "module": "MyNewModule", "type": "StandardModule", "message": "Module 'MyNewModule' created"}
```

### `vba_delete_module`

Delete a VBA module. Document modules (Sheet/ThisWorkbook) cannot be deleted — only their code can be cleared via `vba_write_module`.

```
→ vba_delete_module("C:/Projects/MyWorkbook.xlsm", "OldModule")

{"success": true, "message": "Module 'OldModule' deleted"}
```

### `vba_backup`

Export all VBA modules to text files with a JSON manifest.

```
→ vba_backup("C:/Projects/MyWorkbook.xlsm")

{
  "success": true,
  "backup_dir": "C:/Projects/VBA_Backup/20260324_120000",
  "file_count": 8,
  "files": [
    {"name": "Module1", "type": "StandardModule", "file": "...", "lines": 342},
    ...
  ]
}
```

### `vba_list_workbooks`

Find all `.xlsm` files in a directory.

```
→ vba_list_workbooks("C:/Projects", recursive=True)

[
  {"path": "C:/Projects/MyWorkbook.xlsm", "name": "MyWorkbook.xlsm", "size_kb": 245.3, "modified": "2026-03-24 12:00:00"},
  ...
]
```

---

## Usage Examples

### Refactoring VBA Code

> "Read Module1 from MyWorkbook.xlsm, refactor the `CalculateTotal` function to handle edge cases, and write it back."

Claude will:
1. Call `vba_read_module` to get the current code
2. Analyze and refactor the function
3. Call `vba_write_module` to save the updated code

### Code Review

> "Read all VBA code from MyWorkbook.xlsm and review it for potential bugs, security issues, and best practices."

Claude will:
1. Call `vba_read_all` to get the complete codebase
2. Provide a comprehensive code review

### Migration Assistance

> "Read the VBA code from Legacy.xlsm and help me convert it to Python/JavaScript."

Claude will:
1. Call `vba_read_all` to understand the full VBA codebase
2. Analyze the business logic and data flow
3. Generate equivalent code in the target language

### Batch Operations

> "Find all .xlsm files in C:/Projects, list their modules, and backup everything."

Claude will:
1. Call `vba_list_workbooks` with `recursive=True`
2. Call `vba_list_modules` for each workbook
3. Call `vba_backup` for each workbook

---

## How It Works

```
┌─────────────┐     stdio/MCP      ┌──────────────┐     COM/Win32      ┌─────────┐
│  Claude /    │ ◄────────────────► │  VBA MCP     │ ◄────────────────► │  Excel  │
│  AI Client   │   JSON-RPC 2.0    │  Server      │   pywin32          │  VBA    │
│             │                    │  (server.py) │                    │ Project │
└─────────────┘                    └──────────────┘                    └─────────┘
```

1. **Claude** sends tool calls via MCP (stdio transport)
2. **VBA MCP Server** receives the call, connects to Excel via COM
3. **Excel** provides access to the VBProject object model
4. **Server** reads/writes VBA code and returns results to Claude

---

## Troubleshooting

### "Trust access to the VBA project object model" error

Follow the [Excel Configuration](#excel-configuration) steps above. This is a one-time setting.

### Excel is not responding

If Excel is busy (dialog box open, cell editing mode), COM calls will fail. Make sure Excel is idle before using the tools.

### Module not found

Use `vba_list_modules` first to see available module names. Module names are case-sensitive.

### File path issues

Use forward slashes (`/`) or escaped backslashes (`\\`) in file paths. The server normalizes paths automatically.

### Cannot delete Sheet module

Sheet and ThisWorkbook modules are Document modules — they cannot be deleted from VBProject. Use `vba_write_module` with empty code to clear their contents instead.

---

## Contributing

Contributions are welcome! Here are some ways you can help:

- **Bug reports** — Found an issue? [Open an issue](https://github.com/xiongchenghou/vba-mcp-server/issues)
- **Feature requests** — Have an idea? Let's discuss it
- **Pull requests** — Code improvements, new features, documentation
- **Testing** — Try it with different Excel versions and report results
- **Documentation** — Help improve the docs or add translations

### Development Setup

```bash
git clone https://github.com/xiongchenghou/vba-mcp-server.git
cd vba-mcp-server
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Run the server directly for testing
python server.py
```

### Areas for Improvement

- [ ] UserForm support (read/write `.frm` files with controls)
- [ ] Diff-based editing (modify specific lines instead of full replacement)
- [ ] Excel cell/range read/write (beyond VBA code)
- [ ] Watch mode (detect VBA changes in real-time)
- [ ] Support for multiple simultaneous workbooks
- [ ] macOS support via alternative COM bridges

---

## License

[MIT](LICENSE) — Use it freely, commercially or personally.

---

## Star History

If this tool saves you time, please give it a star! It helps others discover it.

---

<p align="center">
  <b>Stop copy-pasting VBA code. Let AI work with it directly.</b>
</p>
