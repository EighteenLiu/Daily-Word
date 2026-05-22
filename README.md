# 月报生成系统

## 目录结构

```text
01_原始日报/
  中转站/        # 放原始 .doc 日报
  清洁站/
02_转换后日报/
  中转站/        # 自动生成 .docx 日报
  清洁站/
03_月报模板/
  monthly_template.docx
04_输出月报/
scripts/
  convert_script.py
  monthly_generator.py
  run_all.py
requirements.txt
```

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 转换日报

```powershell
python scripts\convert_script.py
```

默认会把 `01_原始日报` 下各站点目录中的 `.doc` 转换到 `02_转换后日报` 的同名站点目录中。

也可以指定转换引擎：

```powershell
python scripts\convert_script.py --engine word
python scripts\convert_script.py --engine libreoffice
```

## 一键执行

```powershell
python scripts\run_all.py --month 2025-01
```

当前已完成 `.doc` 到 `.docx` 转换流程。月报内容抽取和模板填充逻辑需要根据日报格式、月报模板占位符规则继续实现。
