from __future__ import annotations

from pathlib import Path

from reit_excel_auditor.transformer import (
    CustomTemplateLayout,
    FIELD_ALIASES,
    SourceTable,
    STANDARD_TEMPLATE_NAMES,
    build_custom_template_rows,
    headers_similarity,
    is_blank,
    match_source_header,
    normalize_header,
    standard_auto_filter_ref,
    standard_template_names_for,
    to_excel_number,
    to_ratio,
    transform_property,
)


def test_to_ratio_keeps_excel_percent_formatted_values() -> None:
    assert to_ratio(1.006, source_number_format="0.00%") == 1.006


def test_to_ratio_scales_general_whole_percent_values() -> None:
    assert to_ratio(100.13) == 1.0013


def test_to_ratio_scales_percent_string_values() -> None:
    assert to_ratio("100.13%") == 1.0013


def test_to_ratio_preserves_small_general_ratio_behavior() -> None:
    assert to_ratio(0.9521) == 0.9521


def test_non_data_placeholder_values_are_treated_as_blank() -> None:
    assert is_blank(" 文档章节未提及 ")
    assert to_excel_number("文档章节未提及") is None
    assert to_ratio("文档章节未提及") is None


def test_normalize_header_handles_spaces_and_full_width_symbols() -> None:
    assert normalize_header(" REITs 代码（元） ") == "reits代码(元)"


def test_match_source_header_uses_field_aliases() -> None:
    match = match_source_header("资产项目名称", ["REITs代码", "基础设施项目名称"])
    assert match.status == "matched"
    assert match.header == "基础设施项目名称"


def test_match_source_header_uses_common_custom_template_aliases() -> None:
    code_match = match_source_header("REITs代码", ["证券代码", "项目名称"])
    start_match = match_source_header("STARTDATE", ["报告期开始日期", "项目名称"])

    assert code_match.status == "matched"
    assert code_match.header == "证券代码"
    assert start_match.status == "matched"
    assert start_match.header == "报告期开始日期"


def test_template_names_are_loaded_from_release_config() -> None:
    assert STANDARD_TEMPLATE_NAMES["valuation"] == "01-基础资产估值标准模板.xlsx"
    assert standard_template_names_for("valuation") == ["01-基础资产估值标准模板.xlsx"]


def test_field_aliases_keep_default_and_external_config_entries() -> None:
    assert "基础设施项目名称" in FIELD_ALIASES["资产项目名称"]
    assert "基金名称" in FIELD_ALIASES["REITs名称"]


def test_headers_similarity_compares_normalized_sets() -> None:
    reference = ["REITs代码", "资产项目名称", "租金收缴率"]
    current = [" REITs 代码 ", "资产项目名称", "租金收缴率", "额外字段"]
    assert headers_similarity(reference, current) == 0.75


def test_standard_auto_filter_ref_caps_columns_and_uses_output_row_count() -> None:
    assert standard_auto_filter_ref("energy", "A1:K1", header_count=10, row_count=3) == "A1:J4"
    assert standard_auto_filter_ref("traffic", "A1:I16", header_count=9, row_count=0) == "A1:I1"


def test_build_custom_template_rows_reports_missing_and_unused_fields() -> None:
    table = SourceTable(
        path=Path("input.xlsx"),
        sheet_name="Sheet1",
        header_row=1,
        headers=["REITs代码", "基础设施项目名称", "租金收缴率", "未使用字段"],
        rows=[
            {
                normalize_header("REITs代码"): "180301.SZ",
                normalize_header("基础设施项目名称"): "项目A",
                normalize_header("租金收缴率"): 1.006,
                "__format__:" + normalize_header("租金收缴率"): "0.00%",
                normalize_header("未使用字段"): "unused",
            }
        ],
    )
    layout = CustomTemplateLayout(
        path=Path("template.xlsx"),
        header_row=1,
        headers=["REITs代码", "资产项目名称", "租金收缴率", "模板缺失字段"],
        data_number_formats=["General", "General", "0.00%", "General"],
    )
    warnings: list[str] = []

    rows = build_custom_template_rows(table, layout, warnings)

    assert rows == [["180301.SZ", "项目A", 1.006, None]]
    assert any("模板缺失字段" in warning for warning in warnings)
    assert any("未使用字段" in warning for warning in warnings)


def test_transform_property_blanks_non_data_placeholders() -> None:
    table = SourceTable(
        path=Path("property.xlsx"),
        sheet_name="Sheet1",
        header_row=1,
        headers=["REITs代码", "基础设施项目名称", "主配套资产类别", "租金单价(单位:元/月/平方米or元/月/个)", "数据来源"],
        rows=[
            {
                normalize_header("REITs代码"): "180301.SZ",
                normalize_header("基础设施项目名称"): "项目A",
                normalize_header("主配套资产类别"): "主要资产",
                normalize_header("租金单价(单位:元/月/平方米or元/月/个)"): "文档章节未提及",
                normalize_header("数据来源"): "文档章节未提及",
            }
        ],
    )

    rows = transform_property(table, {}, [])

    assert rows[0]["租金单价(元/月/平方米)"] is None
    assert rows[0]["数据来源"] is None
