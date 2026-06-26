# word-mcp-live-cheemscheems 使用指引

## 两种模式的选择

### 1. 跨平台 python-docx 工具（文件必须已关闭、不在 Word 中打开）

**适用场景：**
- 创建新文档
- 批量处理已保存的文件
- Word 未运行时

**特点：**
- 三平台通用（Windows/macOS/Linux）
- 不需要 Word 运行
- 读写已保存的 .docx 文件
- **文件不能被 Word 打开**，否则会报文件锁定错误

**典型工具：** `create_document`、`add_paragraph`、`add_table`、`search_and_replace`、`add_comment`、`get_document_text` 等。

### 2. COM/JXA 实时编辑工具（文件必须在 Word 中打开）

**适用场景：**
- 文档已在 Word 中打开，需要实时修改
- 需要在 Word 里看到变化立即生效
- python-docx 写入因文件锁失败时，应自动切换到此模式

**特点：**
- 工具名以 `word_live_` 开头
- 需要 Word 正在运行且文档已打开
- 每步操作都是单独的 Ctrl+Z 撤销项
- Windows 可用全部 44 个，macOS 可用 40 个

**典型工具：** `word_live_insert_text`、`word_live_replace_text`、`word_live_add_table`、`word_live_format_text` 等。

## 常见问题处理

### 文件被锁定（PermissionError / PackageNotFoundError）

```
错误："File locked (probably open in Word)"
```

**原因：** 用户正在 Word 中查看文档，python-docx 工具无法写入。

**处理步骤：**
1. 确认文档是否已在 Word 中打开
2. 如果已打开，改用对应的 `word_live_*` 工具
3. 例如 `add_table()` → `word_live_add_table()`

### 保存失败

如果 python-docx 工具返回保存失败：
1. 检查文件是否被 Word 锁定
2. 如果是，切换到 live 工具
3. 如果 live 工具也失败，建议用户先保存并关闭文档再重试

### 实时编辑工具不可用（仅 Windows/macOS）

在 Linux 上 `word_live_*` 工具不可用。此时只能：
1. 关闭 Word 中的文档
2. 使用 python-docx 工具操作
3. 告知用户需要在 Windows/macOS 上才能实时编辑

## 推荐工作流

### 用户已打开文档 + 需要大幅修改

```
1. 先用 word_live_get_text 获取当前内容
2. 分析需要修改的部分
3. 用 word_live_replace_text / word_live_insert_text 等实时修改
4. 修改完成后用 word_live_save 保存
```

### 用户未打开文档 + 需要小幅修改

```
1. 用 get_document_text 读取内容
2. 用 add_paragraph / search_and_replace 等修改
3. 自动保存到文件
```

### 大改动前备份

```python
# 先备份
backup_document(filename="文档.docx", note="大改动前备份")
# 再修改
...
```
