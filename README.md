# Daily-Word 日报生成工具

根据垃圾分类检查台账自动生成各街道检查日报。工具支持 `.xls` / `.xlsx` 台账、docxtpl / Jinja2 格式 Word 模板、图片自动匹配插入、图片压缩和批量输出。

GitHub 地址：

https://github.com/EighteenLiu/Daily-Word.git

## 功能

- 从台账中读取街道、点位、指标、具体问题和问题照片。
- 按街道批量生成日报 `.docx`。
- 可在前端同时生成“每日汇总情况”和“垃圾分类工作日报”两个日汇总报告。
- 支持 Jinja2 Word 模板，保留模板页边距、字体、段落和图片位置。
- 支持从 `.xls` 中提取原始嵌入图片，再在日报中插入压缩副本。
- 支持图片压缩等级：标准、轻度、强力、不压缩。
- 使用 Microsoft Excel COM 将 `.xls` 转换为 `.xlsx`，尽量保留台账格式、合并单元格和图片锚点。

## 目录说明

```text
Daily-Word/
  input/                         # 台账、模板建议放这里
  output/                        # 默认输出目录
  scripts/                       # 业务代码
  packaging/                     # 打包脚本
  release/日报生成工具/           # 打包后的可迁移目录
  启动日报生成工具.bat
  使用手册.txt
```

## 直接使用源码运行

双击：

```text
启动日报生成工具.bat
```

如果是在源码目录中运行，bat 会优先使用本机 Python 启动：

```text
scripts\daily_report_gui.py
```

如果当前目录或 `release\日报生成工具\` 下已经有 `日报生成工具.exe`，bat 会优先启动 exe。

也可以在命令行启动前端 GUI：

```powershell
python scripts\daily_report_gui.py
```

如果你的电脑同时安装了多个 Python，推荐用：

```powershell
py -3 scripts\daily_report_gui.py
```

如果已经打包完成，也可以从命令行启动发布版前端：

```powershell
.\release\日报生成工具\日报生成工具.exe
```

## GUI 使用步骤

1. 选择台账文件，支持 `.xls` / `.xlsx`。
2. 可选选择“桶外摆日报”Word 文件，支持 `.docx`，也支持 `.doc` 自动转换。
3. 选择街道日报模板 `.docx`。
4. 如需生成日汇总，勾选“同时生成每日汇总和垃圾分类工作日报”，并选择：
   - 垃圾分类日报模板：例如 `input/垃圾分类工作日报_jinja模板.docx`
   - 每日汇总模板：例如 `input/每日汇总情况_jinja模板.docx`
5. 选择图片压缩等级：
   - 标准：推荐，体积和清晰度平衡。
   - 轻度：更清晰，文件更大。
   - 强力：文件更小。
   - 不压缩：插入原图，文件可能很大。
6. 点击“生成日报”。
7. 到 `output/` 查看生成结果；日汇总默认输出到 `output/每日总结/每日总结（日期）/`。

## Excel COM 转换

如果输入是 `.xls`，程序会调用本机 Microsoft Excel 通过 COM 自动转换为 `.xlsx`。目标电脑需要安装并能正常启动 Excel。

如果输入已经是 `.xlsx`，通常不会调用 Excel 转换。

## 打包 exe

在项目根目录运行：

```powershell
python packaging\build_release.py
```

或：

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_release.ps1
```

输出位置：

```text
release\日报生成工具\
release\日报生成工具.zip
```

迁移给别人时，复制整个 `release\日报生成工具` 文件夹或直接发送 `release\日报生成工具.zip`，不要只复制 exe。

## 命令行

批量生成：

```powershell
python scripts\generate_daily_reports.py input\台账.xls --template input\daily_report_jinja_template.docx --overwrite
```

带桶外摆日报汇入：

```powershell
python scripts\generate_daily_reports.py input\台账.xls --template input\daily_report_jinja_template.docx --outside-bucket-file input\桶外摆日报.docx --overwrite
```

同时生成两个日汇总报告：

```powershell
python scripts\generate_daily_reports.py input\台账.xls --template input\daily_report_jinja_template.docx --garbage-summary-template input\垃圾分类工作日报_jinja模板.docx --daily-summary-template input\每日汇总情况_jinja模板.docx --overwrite
```

指定图片压缩：

```powershell
python scripts\generate_daily_reports.py input\台账.xls --image-compression strong
```

## 模板

推荐使用 docxtpl / Jinja2 `.docx` 模板。

常见写法：

```jinja2
{%p for community in communities %}
（{{ community.index_cn }}）{{ community.name }}
1.小区整体情况
存在的问题是：{{ community.overall_problem_summary }}
{%p for image in community.promo_images %}
{{ image }}
{%p endfor %}
{%p endfor %}
```

注意：

- 控制整段显示时用 `{%p if ... %}` / `{%p for ... %}`。
- 控制同一段中的文字时用普通 `{% if ... %}`。
- 图片段落在模板中设置居中，生成后图片才会居中。
- 动态循环出来的段落样式取决于模板中该段自身样式。

## 常见问题

### 运行时弹 Office 激活窗口

程序现在只使用 Excel COM 转换 `.xls`。如果本机 Office 未激活或 Excel 启动时有弹窗，COM 自动化会被拦截。请先处理 Excel 激活、保护视图、文件恢复等弹窗，再重新生成。

### 生成的 Word 很大

在 GUI 中选择“标准”或“强力”图片压缩。原图仍会保留在 `output/extracted_images`，日报中插入的是临时压缩副本。

### `.xls` 转换失败

请关闭所有 Excel 窗口、激活弹窗、保护视图提示和正在编辑的单元格后重试。也可以手动用 Excel 打开 `.xls`，另存为 `.xlsx` 后再选择 `.xlsx` 台账。

### 模板排版没有保留

确认程序使用的是带 Jinja 标签的 `.docx` 模板，并走 docxtpl 渲染路径。旧的“读模板文字后重建 Word”的方式不会完整保留模板格式。
