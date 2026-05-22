# 街道检查日报生成工具

本项目用于将检查台账 `.xls` 转换为 `.xlsx`，提取点位、问题和照片，并按街道生成检查日报。

## 目录结构

```text
Daily-Word/
  input/                          # 模板、原始 xls、中转站日报 doc/docx
    日报模版.docx
    原始台账.xls
    5月10日中转站检查情况.doc
  output/
    转换xlsx/                     # 转换后的 xlsx
    提取docx/                     # 结构化提取后的 docx
    中转站日报/                   # 中转站 .doc 转换后的 docx
    街道日报（日期）/             # 最终各街道日报
  scripts/
    daily_report_gui.py
    generate_daily_reports.py
    extract_xls_structure.py
    split_street_daily_reports.py
    convert_xls_to_xlsx.py
```

## 使用

双击：

```text
启动日报生成工具.bat
```

或命令行：

```powershell
python scripts\generate_daily_reports.py --overwrite
```

程序默认从 `input` 中读取：

- 日报模板：`input/日报模版.docx`
- 原始台账：最新的 `.xls`
- 中转站日报：匹配日期的 `X月X日中转站检查情况.doc` 或 `.docx`

## 中转站日报规则

中转站日报放在 `input` 中，例如：

```text
input/5月10日中转站检查情况.doc
```

程序会根据当前日报日期查找对应文件，并将其中相关街道内容整体放入对应街道日报的“中转站”一级标题下。该一级标题序号会继承前面的一级标题。

如果中转站日报是 `.doc`，程序会自动转换为 `.docx`，转换结果存放在：

```text
output/中转站日报/
```

## 常见问题

- `Permission denied`：目标 Word 文件正在打开，关闭后重新生成。
- Excel COM 启动失败：请确认安装了 Microsoft Excel。
- Word COM 启动失败：请确认安装了 Microsoft Word，或先手动把中转站 `.doc` 转为 `.docx`。

