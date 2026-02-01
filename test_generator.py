import sys
import types
import unittest

if "pandas" not in sys.modules:
    pandas_stub = types.ModuleType("pandas")
    pandas_stub.read_csv = lambda *args, **kwargs: None
    sys.modules["pandas"] = pandas_stub

if "fpdf" not in sys.modules:
    fpdf_stub = types.ModuleType("fpdf")

    class _DummyFPDF:
        def __init__(self, *args, **kwargs):
            self.font_size = float(kwargs.get("size", 0) or 0)

        def set_auto_page_break(self, *args, **kwargs):
            pass

        def set_margins(self, *args, **kwargs):
            pass

        def add_page(self, *args, **kwargs):
            pass

        def add_font(self, *args, **kwargs):
            pass

        def set_font(self, *args, **kwargs):
            size = kwargs.get("size")
            if size is None and len(args) >= 2:
                size = args[1]
            if size is not None:
                self.font_size = float(size)

        def image(self, *args, **kwargs):
            pass

        def text(self, *args, **kwargs):
            pass

        def get_string_width(self, text):
            return float(len(text))

        def output(self, *args, **kwargs):
            pass

    fpdf_stub.FPDF = _DummyFPDF
    sys.modules["fpdf"] = fpdf_stub

from generator import (
    resolve_split_line_gap,
    resolve_split_style,
    should_split_full_name,
)


class DummyPDF:
    def __init__(self, font_size):
        self.font_size = font_size


class SplitNameBehaviourTests(unittest.TestCase):
    def test_should_split_full_name_uses_default_threshold(self):
        config = {}
        self.assertFalse(should_split_full_name("Anna Nowak", config))
        long_name = "Alicja" + " " + "KowalskanowakowskaTrzecia"
        self.assertTrue(should_split_full_name(long_name, config))

    def test_should_split_full_name_honours_custom_threshold(self):
        config = {"split_name_threshold": 10}
        self.assertTrue(should_split_full_name("Verylong firstname", config))
        config = {"split_name_threshold": 40}
        self.assertFalse(should_split_full_name("Firstname Withspace Lastname", config))

    def test_resolve_split_line_gap_defaults_to_font_height(self):
        pdf = DummyPDF(font_size=32)
        gap = resolve_split_line_gap(pdf, {})
        self.assertAlmostEqual(gap, 32 * 0.85)

    def test_resolve_split_line_gap_uses_configured_value(self):
        pdf = DummyPDF(font_size=10)
        gap = resolve_split_line_gap(pdf, {"split_name_line_gap": 18})
        self.assertEqual(gap, 18)

    def test_resolve_split_style_applies_overrides(self):
        font_size, baseline = resolve_split_style(
            48,
            150,
            {"split_name_font_size": 36, "split_name_text_y": 142},
        )
        self.assertEqual(font_size, 36)
        self.assertEqual(baseline, 142)

    def test_resolve_split_style_falls_back_on_invalid_values(self):
        font_size, baseline = resolve_split_style(
            48,
            150,
            {"split_name_font_size": "bad", "split_name_text_y": "oops"},
        )
        self.assertEqual(font_size, 48)
        self.assertEqual(baseline, 150)

    def test_resolve_split_style_supports_missing_baseline(self):
        font_size, baseline = resolve_split_style(
            40,
            None,
            {"split_name_text_y": 160},
        )
        self.assertEqual(font_size, 40)
        self.assertEqual(baseline, 160)


if __name__ == "__main__":
    unittest.main()
