<div align="center">

# word-mcp-live-cheemscheems

**唯一能在 Word 打开时实时编辑文档的 MCP 服务器**

`实时编辑` &middot; `修订模式` &middot; `单步撤销` &middot; `124 个工具` &middot; `跨平台`

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform: Windows + macOS/Linux](https://img.shields.io/badge/platform-Windows%20%2B%20macOS%2FLinux-lightgrey)]()

</div>

---

word-mcp-live-cheemscheems 让任何支持 [MCP](https://modelcontextprotocol.io/) 的 AI 助手获得对 Microsoft Word 的完全控制。打开文档，告诉 AI 你需要什么，然后看着它一步步完成——格式调整、修订标记、批注、全部实时生效。

## 核心特性

- **实时编辑** — 在 Word 打开的文档上直接编辑，无需保存-关闭-重新打开
- **完整撤销** — 每个 AI 操作对应一个 Ctrl+Z，犯错只需撤销
- **原生修订模式** — 真正的 Word 修订，不是 XML 模拟
- **线程化批注** — 像人类审阅者一样添加、回复、解决和删除批注
- **布局诊断** — 检测格式化问题，在打印前发现隐患
- **公式与交叉引用** — 插入数学公式和自动更新的引用
- **自动备份** — 每 5 分钟自动备份 + 破坏性操作前主动备份，保存在 `_backup/` 文件夹，最多保留 5 份
- **路径沙箱** — 可选 `MCP_ALLOWED_DIR` 环境变量，限制文件访问到指定目录
- **COM 超时保护** — 长时间运行的 COM 操作（替换、保存等）配置超时，防止远程或不稳定的 Word 连接导致服务器挂起
- **安全加固** — 已修复全部 10 项安全审计发现（移除虚假签名、修复 AppleScript 注入、路径遍历防护、消除可预测临时文件等）

## 快速开始

```bash
pip install word-mcp-live-cheemscheems
```

或从源码安装：

```bash
git clone https://github.com/cheemscheems/word-mcp-live.git
cd word-mcp-live
pip install -e .
```

## 客户端配置

<details open>
<summary><b>Claude Desktop</b></summary>

添加到 `claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live-cheemscheems"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Claude Code</b></summary>

添加到 `.mcp.json`：

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live-cheemscheems"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

</details>

<details>
<summary><b>Cursor / VS Code / Windsurf</b></summary>

配置方式与其他客户端类似，详情请参阅[英文文档](README.md#client-installation)。

</details>

> **`MCP_AUTHOR`** 设置修订和批注中的作者名称（默认 `"Author"`）。**`MCP_AUTHOR_INITIALS`** 设置批注缩写。

## 两种工作模式

| | 跨平台模式 | 实时编辑模式 |
|---|---|---|
| **功能** | 创建和编辑已保存的 .docx 文件 | 在 Word 打开时实时编辑 |
| **平台** | Windows、macOS、Linux | Windows (COM) 和 macOS (JXA) |
| **撤销** | 文件级保存 | 单步 Ctrl+Z（Windows）；单步撤销（macOS） |
| **适用场景** | 批量处理、文档生成 | 交互式编辑、格式调整、审阅 |

两种模式协同工作，AI 会根据任务自动选择合适的方式。

## 配置选项

| 变量 | 默认值 | 说明 |
|----------|---------|------|
| `MCP_AUTHOR` | `"Author"` | 修订和批注的作者名称 |
| `MCP_AUTHOR_INITIALS` | `""` | 批注作者缩写 |
| `MCP_TRANSPORT` | `stdio` | 传输类型：`stdio`、`sse` 或 `streamable-http` |
| `MCP_HOST` | `127.0.0.1` | 绑定地址（SSE/HTTP 传输；远程访问需设为 `0.0.0.0`） |
| `MCP_PORT` | `8000` | 绑定端口（SSE/HTTP 传输） |
| `MCP_ALLOWED_DIR` | *(无)* | 路径沙箱：限制文件访问到此目录及其子目录 |
| `MCP_MAX_BACKUPS` | `5` | 每个文档的最大自动备份数，设为 `0` 表示不限制 |
| `WORD_MCP_LIVE_API_KEY` | *(必填)* | HTTP/SSE 传输的 Bearer Token 鉴权。**HTTP/SSE 模式必须设置** |
| `WORD_MCP_LIVE_INSECURE` | *(无)* | 设为 `true` 可关闭鉴权（仅限本地/开发环境，远程部署禁止使用） |

## 安全改进

本项目在初始安全审计后进行了全面加固，修复了全部 10 项发现：

| 发现 | 严重性 | 修复内容 |
|------|--------|----------|
| 虚假密码保护 | 🔴 严重 | 移除 `.protection` 侧边文件，使用真正 OOXML `w:documentProtection` |
| 虚假数字签名 | 🔴 严重 | 移除虚假加密，替换为真正的 Word 签名行 + 正式文本签名区块 |
| AppleScript 注入 | 🟠 高危 | 增强 `_escape_as()` 转义，添加参数类型验证 |
| LibreOffice 子进程 | 🟠 高危 | 添加输出目录白名单验证和路径遍历检测 |
| 默认绑定 `0.0.0.0` | 🟠 高危 | 改为 `127.0.0.1`，远程部署需显式设置 |
| 无路径沙箱 | 🟡 中危 | 新增 `validate_path()` 集中验证 + 可选 `MCP_ALLOWED_DIR` |
| 屏幕捕获暴露 | 🟡 中危 | 使用安全临时文件，返回值包含用户提示 |
| 可预测临时文件 | 🔵 低危 | 改用 `tempfile.mkstemp()` / `NamedTemporaryFile` |
| 缺乏速率限制 | 🔵 低危 | 添加替换操作时间预算（30 秒）和数量上限（5 万次） |
| 文件损坏风险 | 🔵 低危 | 加密/解密改用临时文件 + 原子替换，移除脆弱的数据恢复逻辑 |

## 工具参考

**124 个工具**，详情请参阅 [TOOLS.md](TOOLS.md)。

| 分类 | 数量 |
|----------|-------|
| 跨平台工具 (python-docx) | 80 |
| Windows 实时工具 (COM) | 44 |
| macOS 实时工具 (JXA) | 40（共 44 个实时工具） |

## 系统要求

- **Python 3.11+**
- `python-docx`、`fastmcp`、`msoffcrypto-tool`（自动安装）
- **Windows 实时工具：** Windows 10/11 + Microsoft Word + `pywin32`（自动安装）
- **macOS 实时工具：** macOS + Microsoft Word for Mac（使用内置 JXA，无需额外依赖）

> 跨平台工具无需安装 Word，仅需 python-docx。

## 隐私

本服务器完全在本地运行。不会收集、传输或存储任何数据。详见 [PRIVACY.md](PRIVACY.md)。

## 许可

MIT License — 详见 [LICENSE](LICENSE)。

## 免责声明

> **⚠️ 本项目按现状提供，不提供任何明示或暗示的保证，包括但不限于对适销性或特定用途适用性的保证。**
>
> 虽然初始安全审计中发现的问题已得到修复，但本项目**尚未经过充分的生产环境测试**。使用本软件修改 Word 文档存在固有风险，包括但不限于：
>
> - **数据丢失或损坏** — 文件操作过程中可能因意外错误导致
> - **文档格式问题** — 自动编辑可能导致格式异常
> - **兼容性问题** — 与特定 Word 版本或文档功能可能存在兼容性问题
>
> **在使用本工具之前，请始终对重要文档进行备份。** 自动备份功能（保存在 `_backup/` 文件夹中）提供了一定程度的安全保障，但不应将其作为唯一的备份策略。
>
> 使用本软件即表示您理解并接受这些风险。
