# Word 批量内容提取到 Excel 工具

## 功能简介
本工具可批量读取指定文件夹下的 Word（.docx）文件，按配置文件中的关键字提取内容，自动识别并处理各种对号/空方块符号，仅保留对号后的内容，最终写入 Excel 或 CSV 文件。支持符号映射、灵活关键字、内容截取、自动备份、配置外置等。

## 使用方法

### 1. 环境准备
- 安装 Python 3.7 及以上
- 安装依赖库：
  ```bash
  pip install pandas python-docx openpyxl
  ```

### 2. 配置文件
- 编辑 `config.json`，配置如下内容：
  - `source_folder`：Word 源文件目录
  - `backup_folder`：处理后 Word 文件备份目录
  - `excel_path`：输出 Excel/CSV 路径
  - `keywords`：要提取的字段名（与 Word 文件内容一致）
  - `symbol_maps`：对号/空方块等符号映射（支持 Wingdings 字体）
  - `tick_symbols`：所有视为“对号”的符号（支持 Unicode 转义）
  - `empty_box`：空方块符号

### 3. 使用步骤
1. 将待处理的 Word 文件（.docx）放入 `source_folder` 目录。
2. 配置好 `config.json`。
3. 运行主脚本：
   ```bash
   python word_to_excel.py
   ```
4. 处理完成后，结果文件在 `结果` 文件夹下，原 Word 文件自动移至 `backup_folder`。

### 4. 结果说明
- 每个 Word 文件对应 Excel/CSV 中一行，字段为 `keywords` 配置。
- 字段内有☑等选填项时只保留第一个“☑”或“✔”及其后所有选填项内容，丢弃不选内容如：“☐”。
- 没有对号时保留原文。
- 所有符号映射、对号、空方块均可在 `config.json` 灵活扩展。

### 5. 常见问题
- 如遇符号未被识别，请根据调试输出补全 `config.json` 的 `tick_symbols` 或 `symbol_maps`。
- 支持多种对号、空方块、√、✔、☑、、ü 等符号。
- 仅支持 .docx 格式，暂不支持图片、控件等非文本内容。

---

如需定制符号、关键字或有特殊需求，请联系开发者或继续反馈。
