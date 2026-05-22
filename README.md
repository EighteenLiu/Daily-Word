# 街道检查日报生成工具

本项目用于将检查台账 `.xls` 转换为 `.xlsx`，提取点位、问题和照片，并生成街道日报、垃圾分类工作日总结和每日汇总情况。

## 目录结构

```text
Daily-Word/
  input/                    # 模板、原始 xls、中转站日报
  output/
    转换xlsx/               # 转换后的 xlsx
    提取docx/               # 结构化提取后的 docx
    中转站日报/             # 中转站 .doc 转换后的 docx
    街道日报（日期）/       # 各街道日报
    每日总结（日期）/       # 两个每日总结文件
  scripts/
```

## 窗口使用

优先双击新版 exe：

```text
dist/街道日报生成工具_每日总结版.exe
```

界面中依次选择：

- 日报模板
- 垃圾分类日总结模板
- 每日汇总模板
- 原始 XLS
- 中转站日报，可留空
- 日期，可留空自动识别
- 输出目录，可留空使用默认目录

点击“开始生成”后，会同时生成街道日报和已上传模板对应的每日总结。

## 每日总结规则

垃圾分类日总结只替换“二、区级检查居住小区/平房胡同情况”部分：按街道生成二级标题，街道下按点位写入“点位名存在的问题是：问题内容”。

每日汇总情况只更新开头各街道问题数量，例如“德胜8个，”。“今日市级检查情况更新：”及其后内容保持模板原样。

## 命令行备用

```powershell
python scripts\generate_daily_reports.py input\xxx.xls --template input\日报模版.docx --garbage-summary-template input\垃圾分类工作日报.docx --daily-summary-template input\每日汇总情况.docx --overwrite
```

## 注意

- 生成 `.xls -> .xlsx` 和 `.doc -> .docx` 需要本机安装 Microsoft Excel 和 Microsoft Word。
- 如果出现 `Permission denied`，通常是目标 Word 文件或 exe 正在打开，关闭后重新生成。
- 输出文件默认在 `output` 下，街道日报和每日总结会分目录保存。
