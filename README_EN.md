# REITs Excel Auditor

[中文说明](README.md)

REITs Excel Auditor is a Windows desktop tool for REITs Excel cleanup, conversion, and annual update workflows. The original version converts unaudited REITs workbooks into standardized audit-ready templates. The current version keeps that workflow and adds annual cash-flow update support for property, concession, future cash-flow, and review-summary outputs.

All normal conversion and annual-update processing runs locally. External services are used only when you explicitly enable AI standardization or cloud OCR.

## Features

- Detects and converts the five built-in unaudited REITs workbook types.
- Allows manual table-type selection.
- Supports folder-based batch conversion.
- Supports a helper metadata workbook for REITs name, listing date, announcement date, start date, and end date.
- Supports custom output templates.
- Optionally generates a Type 4 processed workbook for main/supporting asset area repair.
- Supports annual cash-flow update outputs for property, concession, future cash-flow, and review checklists.
- Uses compact annual-update output by default to reduce intermediate files.

## Quick Start

```powershell
python -m pip install -r requirements.txt
python reit_excel_auditor\app.py
```

Optional local OCR dependencies:

```powershell
python -m pip install -r requirements-ocr.txt
```

## Build a Windows Executable

```powershell
.\build_exe.ps1
```

The executable is generated at:

```text
dist\REITsExcelAuditor.exe
```

To include the local OCR engine:

```powershell
.\build_exe.ps1 -WithOCR
```

## Desktop Workflow

1. Choose an input Excel file or folder.
2. Choose an output folder.
3. Select a conversion mode.
4. Optionally choose the metadata workbook.
5. Optionally enable the Type 4 processed output.
6. Optionally choose a custom output template.
7. For annual updates, choose the annual-update workspace and output folder.
8. Start conversion and review the generated workbooks.

## Built-In Conversion Types

| Type | Input Workbook | Output |
| --- | --- | --- |
| Type 1 | Basic asset valuation workbook | Basic asset valuation |
| Type 2 | Traffic infrastructure operation indicators workbook | Expressway operation data |
| Type 3 | Basic asset financial indicators workbook or prospectus asset financial indicators workbook | Assets, liabilities, revenue, and cost |
| Type 4 | Basic asset operation indicators workbook | Property operation data, with optional processed output |
| Type 5 | Energy infrastructure operation indicators workbook | Energy operation data |

## Metadata Workbook

The metadata helper template is:

```text
examples\自动审核补全信息表模板.xlsx
```

It can fill `REITs代码`, `REITs名称`, `上市日期`, `公告日期`, `开始日期`, and `结束日期` when the source workbook does not provide them.

## Custom Template Output

Custom template output is useful when the input workbook is not one of the built-in Type 1-5 formats, but you already have an Excel template that defines the desired output columns and style. The tool reads the template header, column order, formats, and formulas, then maps source fields into the template.

Custom template mode does not run the dedicated Type 4 area-repair logic. Use the built-in Type 4 flow when that repair is needed.

## Annual Cash-Flow Update

The annual-update workflow is designed for yearly REITs updates based on last year's official workbooks, current-year helper data, annual report PDFs, OCR/AI results, and internal standard templates.

Recommended workspace structure:

```text
annual_workspace\
  01_last_checked_workbooks\
    property.xlsx
    concession.xlsx
  02_helper_data\
    年度更新_统一补充大表.xlsx
    年度更新_项目别名映射表.xlsx
    年度更新_基金净资产与折旧摊销参考表.xlsx
  03_annual_report_pdf\
    fund_a_annual_report.pdf
  04_ocr_sources\
    fund_a_cashflow.png
```

Command example:

```powershell
python -m reit_excel_auditor.app --annual-update ".\annual_workspace" --output-dir ".\output"
```

If the unified helper workbook is already prepared and OCR should be skipped:

```powershell
python -m reit_excel_auditor.app --annual-update ".\annual_workspace" --annual-standard-input ".\年度更新_统一补充大表.xlsx" --annual-max-ocr-pages -1 --output-dir ".\output"
```

Compact annual-update outputs usually include:

| File | Description |
| --- | --- |
| `*_自动更新.xlsx` | Updated property or concession workbook. |
| `年度更新_未来现金流汇总表.xlsx` | Future cash-flow summary workbook. |
| `年度更新_结果汇总与复核清单.xlsx` | Combined review workbook containing standardized input, checklist, update plan, field differences, and output comparison. |

Use `--annual-detailed-output-files` if you need the older multi-file intermediate output style.

## Templates

User-fillable blank templates:

```text
examples\
examples\annual_update_helper_templates\
```

Internal standard templates:

```text
standard_templates\excel_conversion\
standard_templates\annual_update\
```

Annual output formatting is controlled by the desensitized templates in `standard_templates\annual_update\`.

## Configuration

| File | Purpose |
| --- | --- |
| `config\table_templates.json` | Maps built-in conversion types to standard templates. |
| `config\field_aliases.json` | Defines source-field aliases. |

Most users do not need to edit these files.

## Before Large-Scale Use

Run:

```powershell
python -m pytest
python scripts\check_private_files.py
```

Make sure real business data, annual report PDFs, OCR images, generated output workbooks, temporary validation folders, local paths, or secrets are not mixed into the working tree.

## Copyright

Copyright (c) 2026 Liu Juncheng. All rights reserved.

This project is intended for REITs Excel formatting, field extraction, annual update, and audit-assistance workflows. Generated outputs should still be reviewed by users. The tool does not provide investment, financial, legal, or audit advice.
