# VBA MCP Server

**AIがExcel VBAコードを直接読み書き — 手動エクスポート/インポート不要。**

[English](README.md) | [中文](README.zh-CN.md) | 日本語

> [MCP（Model Context Protocol）](https://modelcontextprotocol.io/) サーバーです。Claude等のAIアシスタントが、Excelワークブック（`.xlsm`、`.xlsb`、`.xls`、`.xla`、`.xlam`）内のVBAコードに直接アクセス・編集できます。

---

## なぜこのツールが必要？

Excel VBAとAIコーディングアシスタントを両方使っている方なら、この苦痛はご存知でしょう：

1. VBAエディタを開いて、手動でコードをコピー
2. チャットに貼り付けて、AIに修正を依頼
3. AIの回答をコピー
4. VBAエディタに貼り付けて、テストして、また最初から……

**VBA MCP Serverはこのプロセスを完全に排除します。** Claudeが直接VBAモジュールの読み取り・書き込み・作成・削除・バックアップを行えます——通常のソースコードファイルを扱うのと同じように。

---

## 機能

- **VBAに直接アクセス** — エクスポート/インポートなしでVBAコードを読み書き
- **スマートなワークブック処理** — 既に開いているワークブックを自動検出、競合を回避
- **自動保存** — 書き込み後にワークブックを自動保存
- **完全なモジュール管理** — 作成・削除・読み取り・書き込み・一覧・バックアップ
- **ワークブック検索** — ディレクトリツリー内の全 `.xlsm` ファイルを検索
- **マニフェスト付きバックアップ** — 全モジュールをテキストファイルにエクスポート（JSONメタデータ付き）
- **ドキュメントモジュール保護** — Sheet/ThisWorkbookモジュールの誤削除を防止
- **軽量** — 単一Pythonファイル（約300行）、依存関係は2つだけ

---

## 前提条件

| 要件 | 詳細 |
|------|------|
| **OS** | Windows（COM/Win32 APIを使用） |
| **Excel** | Microsoft Excel（ローカルインストール） |
| **Python** | 3.10以上 |
| **信頼設定** | 下記の [Excelの設定](#excelの設定) を参照 |

---

## インストール

### 方法1：pip インストール（推奨）

```bash
pip install vba-mcp-server
```

Claude CodeのMCP設定に追加：

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

### 方法2：Claude Code 手動設定

Claude CodeのMCP設定（`~/.claude.json` またはプロジェクト `.mcp.json`）に追加：

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

### 方法3：Claude Desktop

Claude Desktopの設定ファイル（`%APPDATA%/Claude/claude_desktop_config.json`）に追加：

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

### 方法4：ソースからインストール

```bash
# リポジトリをクローン
git clone https://github.com/xiongchenghou/vba-mcp-server.git
cd vba-mcp-server

# 仮想環境を作成（推奨）
python -m venv .venv
.venv\Scripts\activate

# 依存関係をインストール
pip install -r requirements.txt
```

---

## Excelの設定

**一度だけの設定** — ExcelでVBAへのプログラムアクセスを許可する必要があります：

1. Excelを開く
2. **ファイル** > **オプション** > **トラストセンター** > **トラストセンターの設定** に移動
3. **マクロの設定** をクリック
4. **「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」** にチェックを入れる
5. **OK** をクリック

> この設定を有効にしないと、すべてのツールがヒントメッセージ付きのエラーを返します。

---

## 利用可能なツール

### `vba_list_modules`

ワークブック内の全VBAモジュールをメタデータ付きで一覧表示。

```
→ vba_list_modules("C:/Projects/MyWorkbook.xlsm")

[
  {"name": "Module1", "type": "StandardModule", "line_count": 342},
  {"name": "Sheet1",  "type": "Document",       "line_count": 15},
  {"name": "MyClass", "type": "ClassModule",     "line_count": 89}
]
```

### `vba_read_module`

指定モジュールのソースコードを読み取り。

```
→ vba_read_module("C:/Projects/MyWorkbook.xlsm", "Module1")

' === Module: Module1 (StandardModule) ===
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub
```

### `vba_read_all`

全モジュールを一括読み取り — コード全体のレビューに最適。

```
→ vba_read_all("C:/Projects/MyWorkbook.xlsm")

============================================================
' Module: Module1
' Type: StandardModule
' Lines: 342
============================================================
[完全なソースコード...]
```

### `vba_write_module`

モジュールのコードを書き込みまたは置換。ワークブックは自動保存されます。モジュールが存在しない場合は自動作成。

```
→ vba_write_module("C:/Projects/MyWorkbook.xlsm", "Module1", "Sub Test()\n    MsgBox \"Updated!\"\nEnd Sub")

{"success": true, "module": "Module1", "lines_written": 3, "message": "Module 'Module1' updated (3 lines)"}
```

| パラメータ | 型 | デフォルト | 説明 |
|-----------|------|-----------|------|
| `file_path` | string | 必須 | ワークブックのパス |
| `module_name` | string | 必須 | 書き込むモジュール名 |
| `code` | string | 必須 | VBAソースコード |
| `create_if_missing` | bool | `true` | モジュールが存在しない場合に自動作成 |

### `vba_create_module`

新しいVBAモジュールを作成（標準モジュールまたはクラスモジュール）。

```
→ vba_create_module("C:/Projects/MyWorkbook.xlsm", "MyNewModule", "Option Explicit", "standard")

{"success": true, "module": "MyNewModule", "type": "StandardModule", "message": "Module 'MyNewModule' created"}
```

### `vba_delete_module`

VBAモジュールを削除。ドキュメントモジュール（Sheet/ThisWorkbook）は削除できません — `vba_write_module` で空のコードを渡してクリアしてください。

```
→ vba_delete_module("C:/Projects/MyWorkbook.xlsm", "OldModule")

{"success": true, "message": "Module 'OldModule' deleted"}
```

### `vba_backup`

全VBAモジュールをテキストファイルにエクスポート（JSONマニフェスト付き）。

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

ディレクトリ内の全 `.xlsm` ファイルを検索。

```
→ vba_list_workbooks("C:/Projects", recursive=True)

[
  {"path": "C:/Projects/MyWorkbook.xlsm", "name": "MyWorkbook.xlsm", "size_kb": 245.3, "modified": "2026-03-24 12:00:00"}
]
```

---

## 使用例

### VBAコードのリファクタリング

> 「MyWorkbook.xlsmのModule1を読み取って、`CalculateTotal`関数をエッジケースに対応するようリファクタリングして、書き戻して。」

Claudeが自動的に：モジュール読み取り → 分析・リファクタリング → 書き込み・保存

### コードレビュー

> 「MyWorkbook.xlsmの全VBAコードを読み取って、バグ・セキュリティ問題・ベストプラクティスの観点でレビューして。」

### コード移行

> 「Legacy.xlsmのVBAコードを読み取って、Python/JavaScriptに変換する手助けをして。」

### 一括操作

> 「C:/Projects配下の全.xlsmファイルを見つけて、モジュールを一覧表示して、全部バックアップして。」

---

## 仕組み

```
┌─────────────┐     stdio/MCP      ┌──────────────┐     COM/Win32      ┌─────────┐
│  Claude /    │ ◄────────────────► │  VBA MCP     │ ◄────────────────► │  Excel  │
│  AIクライアント│   JSON-RPC 2.0    │  Server      │   pywin32          │  VBA    │
│             │                    │  (server.py) │                    │ Project │
└─────────────┘                    └──────────────┘                    └─────────┘
```

1. **Claude** がMCP経由でツール呼び出しを送信（stdio転送）
2. **VBA MCP Server** が呼び出しを受信し、COM経由でExcelに接続
3. **Excel** がVBProjectオブジェクトモデルへのアクセスを提供
4. **Server** がVBAコードを読み書きし、結果をClaudeに返却

> **クラウドにデプロイできないのはなぜ？** サーバーはWindows COM経由でローカルのExcelプロセスと通信する必要があるためです。ただし `pip install` で簡単にインストールでき、他のMCPサーバーと同じ手軽さで使えます。

---

## トラブルシューティング

### 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」エラー

上記の [Excelの設定](#excelの設定) の手順に従ってください。一度だけの設定です。

### Excelが応答しない

Excelがビジー状態（ダイアログ表示中、セル編集中）の場合、COM呼び出しは失敗します。Excelがアイドル状態であることを確認してください。

### モジュールが見つからない

まず `vba_list_modules` でモジュール名を確認してください。モジュール名は大文字小文字を区別します。

### ファイルパスの問題

スラッシュ（`/`）またはエスケープされたバックスラッシュ（`\\`）を使用してください。サーバーは自動的にパスを正規化します。

### Sheetモジュールを削除できない

SheetとThisWorkbookはドキュメントモジュールであり、VBProjectから削除できません。代わりに `vba_write_module` で空のコードを渡してクリアしてください。

---

## コントリビュート

あらゆる貢献を歓迎します！

- **バグ報告** — 問題を発見したら [Issue を作成](https://github.com/xiongchenghou/vba-mcp-server/issues)
- **機能リクエスト** — アイデアをお持ちですか？ぜひ議論しましょう
- **Pull Request** — コード改善、新機能、ドキュメント整備
- **テスト** — 異なるExcelバージョンでテストして結果を報告
- **翻訳** — ドキュメントの改善や他言語への翻訳

### 改善予定

- [ ] UserFormサポート（`.frm`ファイルとコントロールの読み書き）
- [ ] 差分編集（全体置換ではなく特定行の変更）
- [ ] Excelセル/範囲の読み書き（VBAコード以外も）
- [ ] ウォッチモード（VBA変更のリアルタイム検出）
- [ ] 複数ワークブックの同時操作
- [ ] macOS対応（代替COMブリッジ経由）

---

## ライセンス

[MIT](LICENSE) — 商用・個人問わず自由に使用可能。

---

<p align="center">
  <b>VBAコードのコピペはもう終わり。AIに直接操作させましょう。</b>
</p>
