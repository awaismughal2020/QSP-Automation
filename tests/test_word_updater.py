"""Tests for Word template updater helpers."""

from src.generators.word_updater import (
    build_quarter_replacement_pairs,
    discover_quarter_references,
    format_report_date_dutch,
    _collapse_xml_runs_to_single_value,
    _find_fragmented_number_span,
)
from datetime import datetime


class TestQuarterDiscovery:
    def test_discovers_multiple_stale_quarters(self):
        text = "Portfolio highlights – Q1 2025. Footer Q2 2025. Target Q1 2026."
        found = discover_quarter_references(text, exclude="Q1 2026")
        assert "Q1 2025" in found
        assert "Q2 2025" in found
        assert "Q1 2026" not in found

    def test_build_replacement_pairs_for_stale_quarters(self):
        pairs = build_quarter_replacement_pairs(
            ["Q1 2025", "Q2 2025"], "Q1 2026"
        )
        assert ("Q1 2025", "Q1 2026") in pairs
        assert ("Q2 2025", "Q1 2026") in pairs
        assert ("first quarter", "first quarter") not in pairs or any(
            p[0] == "first quarter" for p in pairs
        )


class TestReportDate:
    def test_dutch_format(self):
        assert format_report_date_dutch(datetime(2026, 6, 7)) == "7 juni 2026"


class TestFragmentedXmlNumbers:
    def test_collapse_runs_to_single_value(self):
        xml = (
            '<w:t>3,</w:t><w:t>200</w:t><w:t>.6</w:t>'
        )
        out = _collapse_xml_runs_to_single_value(xml, "3,332.7")
        assert ">3,332.7</w:t>" in out
        assert ">200</w:t>" not in out

    def test_find_fragmented_number_span(self):
        content = (
            'prefix<w:t>€</w:t><w:t>3,</w:t><w:t>200</w:t><w:t>.6</w:t>'
            '<w:t>k</w:t>suffix'
        )
        span = _find_fragmented_number_span(content, 3200.6)
        assert span is not None
        assert "3," in span
