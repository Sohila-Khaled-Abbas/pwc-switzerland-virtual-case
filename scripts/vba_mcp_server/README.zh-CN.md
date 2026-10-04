# VBA MCP Server

**让 AI 直接读写 Excel VBA 代码 —— 告别手动导出/导入。**

[English](README.md) | 中文 | [日本語](README.ja.md)

> 一个 [MCP（Model Context Protocol）](https://modelcontextprotocol.io/) 服务器，让 Claude 等 AI 助手能够直接访问和修改 Excel 工作簿（`.xlsm`、`.xlsb`、`.xls`、`.xla`、`.xlam`）中的 VBA 代码。

---

## 为什么需要它？

如果你同时使用 Excel VBA 和 AI 编程助手，你一定经历过这个痛苦的循环：

1. 打开 VBA 编辑器，手动复制代码
2. 粘贴到聊天窗口，请 AI 帮忙修改
3. 复制 AI 的回复
4. 粘贴回 VBA 编辑器，测试，再来一遍……

**VBA MCP Server 彻底消除了这个过程。** Claude 现在可以直接读取、编写、创建、删除和备份 VBA 模块——就像操作普通源代码文件一样。

---

## 功能特性

- **直接访问 VBA** — 原地读写 VBA 代码，无需导出/导入
- **智能工作簿处理** — 自动检测已打开的工作簿，避免冲突
- **自动保存** — 写入后自动保存工作簿
- **完整的模块管理** — 创建、删除、读取、写入、列出和备份模块
- **工作簿发现** — 在目录树中查找所有 `.xlsm` 文件
- **带清单的备份** — 将所有模块导出为文本文件，附带 JSON 元数据
- **文档模块保护** — 防止误删 Sheet/ThisWorkbook 模块
- **轻量级** — 单个 Python 文件（约 300 行），仅需两个依赖

---

## 前置要求

| 要求 | 详情 |
|------|------|
| **操作系统** | Windows（使用 COM/Win32 API） |
| **Excel** | Microsoft Excel（本地安装） |
| **Python** | 3.10 或更高版本 |
| **信任设置** | 参见下方 [Excel 配置](#excel-配置) |

---

## 安装方式

### 方式一：pip 安装（推荐）

```bash
pip install vba-mcp-server
```

然后在 Claude Code 的 MCP 配置中添加：

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

### 方式二：Claude Code 手动配置

添加到你的 Claude Code MCP 设置（`~/.claude.json` 或项目 `.mcp.json`）：

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

### 方式三：Claude Desktop

添加到 Claude Desktop 配置文件（`%APPDATA%/Claude/claude_desktop_config.json`）：

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

### 方式四：从源码安装

```bash
# 克隆仓库
git clone https://github.com/xiongchenghou/vba-mcp-server.git
cd vba-mcp-server

# 创建虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

---

## Excel 配置

**一次性设置** — Excel 必须允许程序化访问 VBA：

1. 打开 Excel
2. 进入 **文件** > **选项** > **信任中心** > **信任中心设置**
3. 点击 **宏设置**
4. 勾选 **"信任对 VBA 工程对象模型的访问"**
5. 点击 **确定**

> 如果没有开启此设置，所有工具调用都会返回一个包含提示信息的错误。

---

## 可用工具

### `vba_list_modules`

列出工作簿中所有 VBA 模块及其元数据。

```
→ vba_list_modules("C:/Projects/MyWorkbook.xlsm")

[
  {"name": "Module1", "type": "StandardModule", "line_count": 342},
  {"name": "Sheet1",  "type": "Document",       "line_count": 15},
  {"name": "MyClass", "type": "ClassModule",     "line_count": 89}
]
```

### `vba_read_module`

读取指定模块的源代码。

```
→ vba_read_module("C:/Projects/MyWorkbook.xlsm", "Module1")

' === Module: Module1 (StandardModule) ===
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub
```

### `vba_read_all`

一次性读取所有模块 —— 非常适合全量代码审查。

```
→ vba_read_all("C:/Projects/MyWorkbook.xlsm")

============================================================
' Module: Module1
' Type: StandardModule
' Lines: 342
============================================================
[完整源代码...]
```

### `vba_write_module`

写入或替换模块代码。自动保存工作簿。如果模块不存在则自动创建。

```
→ vba_write_module("C:/Projects/MyWorkbook.xlsm", "Module1", "Sub Test()\n    MsgBox \"Updated!\"\nEnd Sub")

{"success": true, "module": "Module1", "lines_written": 3, "message": "Module 'Module1' updated (3 lines)"}
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `file_path` | string | 必填 | 工作簿路径 |
| `module_name` | string | 必填 | 要写入的模块名 |
| `code` | string | 必填 | VBA 源代码 |
| `create_if_missing` | bool | `true` | 模块不存在时是否自动创建 |

### `vba_create_module`

创建新的 VBA 模块（标准模块或类模块）。

```
→ vba_create_module("C:/Projects/MyWorkbook.xlsm", "MyNewModule", "Option Explicit", "standard")

{"success": true, "module": "MyNewModule", "type": "StandardModule", "message": "Module 'MyNewModule' created"}
```

### `vba_delete_module`

删除 VBA 模块。文档模块（Sheet/ThisWorkbook）不能删除——只能通过 `vba_write_module` 清空其代码。

```
→ vba_delete_module("C:/Projects/MyWorkbook.xlsm", "OldModule")

{"success": true, "message": "Module 'OldModule' deleted"}
```

### `vba_backup`

将所有 VBA 模块导出为文本文件，并附带 JSON 清单。

```
→ vba_backup("C:/Projects/MyWorkbook.xlsm")

{
  "success": true,
  "backup_dir": "C:/Projects/VBA_Backup/20260324_120000",
  "file_count": 8,
  "files": [
    {"name": "Module1", "type": "StandardModule", "file": "...", "lines": 342}
  ]
}
```

### `vba_list_workbooks`

在指定目录中查找所有 `.xlsm` 文件。

```
→ vba_list_workbooks("C:/Projects", recursive=True)

[
  {"path": "C:/Projects/MyWorkbook.xlsm", "name": "MyWorkbook.xlsm", "size_kb": 245.3, "modified": "2026-03-24 12:00:00"}
]
```

---

## 使用场景

### 重构 VBA 代码

> "读取 MyWorkbook.xlsm 中的 Module1，重构 `CalculateTotal` 函数以处理边界情况，然后写回去。"

Claude 会自动：读取模块 → 分析重构 → 写入保存

### 代码审查

> "读取 MyWorkbook.xlsm 的所有 VBA 代码，审查潜在的 Bug、安全问题和最佳实践。"

### 代码迁移

> "读取 Legacy.xlsm 的 VBA 代码，帮我转换为 Python/JavaScript。"

### 批量操作

> "查找 C:/Projects 下所有 .xlsm 文件，列出它们的模块，并全部备份。"

---

## 工作原理

```
┌─────────────┐     stdio/MCP      ┌──────────────┐     COM/Win32      ┌─────────┐
│  Claude /    │ ◄────────────────► │  VBA MCP     │ ◄────────────────► │  Excel  │
│  AI 客户端   │   JSON-RPC 2.0    │  Server      │   pywin32          │  VBA    │
│             │                    │  (server.py) │                    │ Project │
└─────────────┘                    └──────────────┘                    └─────────┘
```

1. **Claude** 通过 MCP 发送工具调用（stdio 传输）
2. **VBA MCP Server** 接收调用，通过 COM 连接 Excel
3. **Excel** 提供对 VBProject 对象模型的访问
4. **Server** 读写 VBA 代码并将结果返回给 Claude

> **为什么不能部署到云端？** 因为服务器需要通过 Windows COM 接口与本地 Excel 进程通信。这是 Windows + Excel 的本地操作，无法远程化。但你可以通过 `pip install` 一键安装，跟其他 MCP 一样方便。

---

## 常见问题

### "Trust access to the VBA project object model" 错误

按照上方 [Excel 配置](#excel-配置) 步骤操作。这是一次性设置。

### Excel 无响应

如果 Excel 正忙（弹窗打开、单元格编辑中），COM 调用会失败。确保 Excel 处于空闲状态。

### 找不到模块

先用 `vba_list_modules` 查看可用的模块名。模块名区分大小写。

### 文件路径问题

使用正斜杠（`/`）或转义反斜杠（`\\`）。服务器会自动规范化路径。

### 无法删除 Sheet 模块

Sheet 和 ThisWorkbook 是文档模块，无法从 VBProject 中删除。使用 `vba_write_module` 传入空代码来清空其内容。

---

## 参与贡献

欢迎任何形式的贡献！

- **Bug 报告** — 发现问题？[提交 Issue](https://github.com/xiongchenghou/vba-mcp-server/issues)
- **功能建议** — 有好想法？一起讨论
- **Pull Request** — 代码改进、新功能、文档完善
- **测试** — 在不同 Excel 版本下测试并反馈结果
- **翻译** — 帮助改进文档或添加更多语言版本

### 改进方向

- [ ] UserForm 支持（读写 `.frm` 文件及控件）
- [ ] 差异化编辑（修改特定行而非全量替换）
- [ ] Excel 单元格/区域读写（不仅限于 VBA 代码）
- [ ] 监听模式（实时检测 VBA 变更）
- [ ] 支持同时操作多个工作簿
- [ ] macOS 支持（通过替代 COM 桥接方案）

---

## 许可证

[MIT](LICENSE) — 自由使用，商用/个人均可。

---

<p align="center">
  <b>别再复制粘贴 VBA 代码了。让 AI 直接操作它。</b>
</p>
