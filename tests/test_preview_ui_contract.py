import re
from pathlib import Path
import unittest

from formatter.localization import t


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
LOCALIZATION_PATH = Path(__file__).resolve().parents[1] / "formatter" / "localization.py"


def _hex_to_rgb(value: str) -> tuple[float, float, float]:
    value = value.lstrip("#")
    return tuple(int(value[index : index + 2], 16) / 255 for index in (0, 2, 4))


def _linear_channel(channel: float) -> float:
    if channel <= 0.03928:
        return channel / 12.92
    return ((channel + 0.055) / 1.055) ** 2.4


def _contrast(foreground: str, background: str) -> float:
    fg = _hex_to_rgb(foreground)
    bg = _hex_to_rgb(background)
    fg_luminance = sum(weight * _linear_channel(channel) for weight, channel in zip((0.2126, 0.7152, 0.0722), fg))
    bg_luminance = sum(weight * _linear_channel(channel) for weight, channel in zip((0.2126, 0.7152, 0.0722), bg))
    lighter = max(fg_luminance, bg_luminance)
    darker = min(fg_luminance, bg_luminance)
    return (lighter + 0.05) / (darker + 0.05)


class PreviewUiContractTest(unittest.TestCase):
    def setUp(self):
        self.app_source = APP_PATH.read_text(encoding="utf-8")

    def test_active_preview_route_runs_before_upload_section(self):
        preview_index = self.app_source.index("if _render_active_preview_from_session(ui_language):")
        upload_index = self.app_source.index("report_file = st.file_uploader(")

        self.assertLess(preview_index, upload_index)
        self.assertIn('"latest_original_docx_bytes"', self.app_source)
        self.assertIn('"latest_original_file_name"', self.app_source)

    def test_unrequested_input_files_region_is_not_rendered(self):
        for forbidden in [
            'st.subheader(t(ui_language, "input_files"))',
            "target_report_title",
            "target_report_desc",
            "required_badge",
            "upload-title",
            "upload-desc",
            "upload-badge",
        ]:
            self.assertNotIn(forbidden, self.app_source)

        self.assertIn('t(ui_language, "upload_report")', self.app_source)
        self.assertIn('t(ui_language, "upload_template")', self.app_source)

    def test_preview_route_has_local_top_spacer_before_toolbar(self):
        workspace_index = self.app_source.index("def _render_preview_workspace(")
        spacer_index = self.app_source.index("preview-route-top-spacer", workspace_index)
        toolbar_index = self.app_source.index("header_cols = st.columns", workspace_index)

        self.assertLess(spacer_index, toolbar_index)
        self.assertIn(".preview-route-top-spacer", self.app_source)
        self.assertIn("height: 4rem;", self.app_source)

    def test_preview_controls_and_page_structure_are_present(self):
        for token in [
            "return_analysis_results",
            "return_previous_level",
            "return_to_previous_level_from_results",
            "previous_page",
            "next_page",
            "page_number",
            "zoom",
            "comparison_view_mode",
            "sync_scroll",
            "sync_zoom",
            "original_file",
            "modified_file",
            "current_page_changes",
            "current_page_issues",
            "preview_performance",
            "preview-columns",
            "preview-pane",
            "preview-canvas-container",
        ]:
            self.assertIn(token, self.app_source)

    def test_results_page_back_button_preserves_latest_result(self):
        nav_index = self.app_source.index("result_nav_cols = st.columns")
        button_index = self.app_source.index("return_to_previous_level_from_results", nav_index)
        rerun_index = self.app_source.index("st.rerun()", button_index)
        button_block = self.app_source[button_index:rerun_index]

        self.assertNotIn("_clear_latest_result()", button_block)
        self.assertIn('st.session_state["results_view_active"] = False', button_block)
        self.assertIn('st.session_state["active_preview_mode"] = None', button_block)

    def test_preview_route_preserves_upload_widget_state(self):
        preserve_index = self.app_source.index("_render_preview_upload_state_keepers(ui_language)")
        preview_index = self.app_source.index("if _render_active_preview_from_session(ui_language):")

        self.assertLess(preserve_index, preview_index)
        self.assertIn('key="preview_upload_state_keeper"', self.app_source)
        self.assertIn(".st-key-preview_upload_state_keeper", self.app_source)
        self.assertIn('key="report_file"', self.app_source)
        self.assertIn('key="template_file"', self.app_source)

    def test_targeted_text_color_selectors_are_scoped(self):
        for token in [
            ".st-key-return_to_previous_level_from_results button",
            ".st-key-return_from_preview button",
            ".st-key-preview_next_page button",
            '[class*="st-key-issue_group_open_"][class*="_toggle"] button',
            '[class*="st-key-preview_page_input_"] [data-testid="stWidgetLabel"]',
            ".st-key-preview_show_markers_control label",
            ".st-key-comparison_sync_scroll_control label",
            ".st-key-comparison_sync_zoom_control label",
            '.st-key-comparison_preview_mode [data-testid="stRadio"] label',
            "color: #FFFFFF !important;",
            "-webkit-text-fill-color: #FFFFFF !important;",
            "color: #273044 !important;",
            "-webkit-text-fill-color: #273044 !important;",
        ]:
            self.assertIn(token, self.app_source)

    def test_semantic_colors_meet_contrast_contracts(self):
        variables = dict(re.findall(r"--([a-z-]+):\s*(#[0-9a-fA-F]{6});", self.app_source))
        for name in [
            "page-bg",
            "surface-bg",
            "text-primary",
            "text-secondary",
            "text-inverse",
            "control-bg",
            "control-surface-bg",
            "control-surface-hover-bg",
            "control-surface-active-bg",
            "control-text",
            "control-hover-text",
            "control-border",
            "control-primary-bg",
            "control-primary-text",
            "disabled-bg",
            "disabled-text",
            "accent-soft",
            "accent-text",
        ]:
            self.assertIn(name, variables)

        self.assertGreaterEqual(_contrast(variables["text-primary"], variables["page-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["text-primary"], variables["surface-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["text-inverse"], variables["control-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["control-text"], variables["control-surface-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["control-hover-text"], variables["control-surface-hover-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["control-text"], variables["control-surface-active-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["control-primary-text"], variables["control-primary-bg"]), 4.5)
        self.assertGreaterEqual(_contrast(variables["disabled-text"], variables["disabled-bg"]), 3.0)
        self.assertGreaterEqual(_contrast(variables["accent-text"], variables["accent-soft"]), 4.5)
        for semantic_alias in [
            "--color-text-primary",
            "--color-text-secondary",
            "--color-text-inverse",
            "--color-surface-page",
            "--color-surface-card",
            "--color-surface-control",
            "--color-control-text",
            "--color-control-hover-text",
            "--color-border-default",
            "--color-border-strong",
            "--color-accent",
            "--color-accent-hover",
            "--color-accent-active",
            "--color-disabled-text",
            "--color-disabled-background",
        ]:
            self.assertIn(semantic_alias, self.app_source)

    def test_preview_layout_uses_responsive_grid(self):
        self.assertIn("grid-template-columns: repeat(2, minmax(0, 1fr));", self.app_source)
        self.assertIn("@media (max-width: 1000px)", self.app_source)
        self.assertIn("grid-template-columns: 1fr;", self.app_source)

    def test_control_state_selectors_are_defined(self):
        for selector in [
            ":hover",
            ":active",
            ":focus-visible",
            ":disabled",
            "[data-testid=\"stCheckbox\"]",
            "[data-testid=\"stRadio\"]",
            "[data-testid=\"stSlider\"]",
            "[data-testid=\"stNumberInput\"]",
        ]:
            self.assertIn(selector, self.app_source)

    def test_streamlit_widget_readability_overrides_are_defined(self):
        for selector in [
            'button:not([data-testid="stBaseButton-primary"]):not([kind="primary"])',
            'button[data-testid="stBaseButton-secondary"]',
            '[data-testid="stWidgetLabel"]',
            '[data-testid="stNumberInput"] [data-baseweb="input"]',
            '[data-testid="stNumberInput"] [role="button"]',
        ]:
            self.assertIn(selector, self.app_source)

    def test_main_secondary_buttons_do_not_use_inverse_text_by_default(self):
        button_rule = re.search(
            r'\[data-testid="stAppViewContainer"\] main div\[data-testid="stButton"\] button \{(?P<body>.*?)\n    \}',
            self.app_source,
            re.DOTALL,
        )
        self.assertIsNotNone(button_rule)
        body = button_rule.group("body")
        self.assertIn("background: var(--control-surface-bg);", body)
        self.assertIn("color: var(--control-text);", body)
        self.assertNotIn("background: var(--control-bg);", body)
        self.assertNotIn("color: var(--text-inverse);", body)
        self.assertIn('button[data-testid="stBaseButton-primary"]', self.app_source)

    def test_buttons_and_selectboxes_use_flex_vertical_alignment(self):
        for token in [
            '[data-testid="stAppViewContainer"] main div[data-testid="stButton"] button {',
            "display: inline-flex;",
            "align-items: center;",
            "justify-content: center;",
            "line-height: 1.2;",
            ".st-key-preview_page_not_cached button",
            '[class*="st-key-"][class*="_thumb_"] button',
            '[class*="st-key-"][class*="thumb"] button',
            'section[data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="value-container"]',
            'section[data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="single-value"]',
            'section[data-testid="stSidebar"] div[data-baseweb="select"] [data-baseweb="select-dropdown"]',
            'div[data-baseweb="popover"] div[role="option"]',
        ]:
            self.assertIn(token, self.app_source)

    def test_difference_view_uses_pixel_markers_not_structured_markers(self):
        for token in [
            "page_pixel_changes = items_for_page(getattr(bundle, \"pixel_changes\", ()), modified_page.page_number)",
            "diff_markers = page_pixel_changes if selected_view == \"diff\" else []",
            "marker_visible=False",
            "pixel_diff_active",
            "pixel_diff_unavailable",
        ]:
            self.assertIn(token, self.app_source)

    def test_comparison_view_mode_keeps_original_radio_shape(self):
        self.assertIn("selected_view = st.radio(", self.app_source)
        self.assertIn('key="comparison_preview_mode"', self.app_source)
        self.assertIn("horizontal=True", self.app_source)
        self.assertNotIn("st.segmented_control(", self.app_source)
        self.assertNotIn('width="stretch"', self.app_source)
        self.assertNotIn('[data-testid="stSegmentedControl"]', self.app_source)

    def test_equation_renumbering_defaults_to_safe_off(self):
        self.assertIn('"renumber_equations": False', self.app_source)
        self.assertIn("_has_equation_auto_edit_suspended_notice", self.app_source)
        self.assertIn("equation_auto_edit_suspended_notice", self.app_source)

    def test_page_count_mismatch_blocks_export_ready_state(self):
        self.assertIn("layout_guard_blocked", self.app_source)
        self.assertIn('st.session_state["latest_preview_version"] = None', self.app_source)
        self.assertIn("Export blocked: the modified file changed", self.app_source)

    def test_main_heading_allows_chinese_title_to_wrap_without_clipping(self):
        self.assertIn("line-height: 1.18", self.app_source)
        self.assertIn("overflow: visible", self.app_source)
        self.assertIn("white-space: normal", self.app_source)

    def test_preview_i18n_has_no_missing_translation_fallback_text(self):
        localization_source = LOCALIZATION_PATH.read_text(encoding="utf-8")
        self.assertNotIn("未翻译文本", localization_source)
        self.assertNotIn("Missing translation", localization_source)
        for language in ("en", "zh"):
            for key in (
                "zoom",
                "preview_performance",
                "preview_page_not_cached",
                "return_previous_level",
                "pixel_diff_active",
                "pixel_diff_unavailable",
            ):
                self.assertNotEqual(t(language, key), "")
                self.assertNotEqual(t(language, key), key)

    def test_chinese_return_buttons_cannot_fall_back_to_english(self):
        self.assertEqual(t("zh", "return_analysis_results"), "返回分析结果")
        self.assertEqual(t("zh", "return_previous_level"), "返回上一级")
        self.assertNotIn(t("zh", "return_previous_level"), {"Return previous level", "Back"})
        self.assertNotIn(t("zh", "return_analysis_results"), {"Return to results", "Back"})

    def test_runtime_diagnostics_show_loaded_localization_module(self):
        for token in [
            "localization_path",
            "localization_mtime",
            "return_analysis_results_label",
            "return_previous_level_label",
        ]:
            self.assertIn(token, self.app_source)


if __name__ == "__main__":
    unittest.main()
