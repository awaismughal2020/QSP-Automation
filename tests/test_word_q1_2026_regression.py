import re
import zipfile
from datetime import datetime
from pathlib import Path

from src.generators.word_updater import build_report_values, WordTemplateUpdater


def _visible_text(path):
    parts = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith('.xml') and (
                n == 'word/document.xml' or '/header' in n or '/footer' in n
            ):
                parts.append(z.read(n).decode('utf-8', 'replace'))
    return re.sub(r'<[^>]+>', '', ''.join(parts))


def test_q1_2026_executive_summary(tmp_path):
    tmpl = Path('ref-files/Quarterly QSP - Q1 2026 - Draft.docx')
    ma_values = {
        'gtri': 3400.0, 'gross_rental': 3120.0, 'vacancy_pct': 6.2,
        'vacancy_amount': 210.0, 'maintenance': 260.0,
        'unit_sales_proceeds': 203.5, 'gtri_ltm': 13500.0,
        'gross_rental_ltm': 12800.0, 'maintenance_ltm': 1000.0,
    }
    rv = build_report_values(
        ma_values=ma_values, rent_roll_k=13317.9, rent_roll_units=400,
        units_sold_quarter=2, unit_sales_proceeds=203.5,
        word_template_path=tmpl, report_date=datetime(2026, 4, 15),
    )
    out = tmp_path / 'out.docx'
    WordTemplateUpdater(str(tmpl), str(out)).update_with_python_docx(
        rv, 'Q4 2025', 'Q1 2026'
    )
    text = _visible_text(out)

    # 1. No stale or wrong-year quarter labels
    assert 'Q1 2025' not in text
    assert 'Q2 2025' not in text
    assert 'Q1 2026' in text

    # 2. No double-decimal artifacts
    assert not re.search(r'\d\.\d\.\d', text)

    # 3. Euro sign preserved on KPI figures
    assert re.search(r'€\s?3,?400', text)          # GTRI body keeps euro
    assert 'income:260' not in text                # maintenance must not land in GTRI slot

    # 4. Highlights-box GTRI populates with the NEW GTRI (3,400), not a stale
    #    template value, and stays consistent with the body.
    m = re.search(r'Gross Theoretical rental income:\s*€?\s?([\d,\.]+)', text)
    assert m, 'no highlights GTRI match'
    assert m.group(1).replace(',', '').startswith('3400'), m.group(1)


def _set_highlights_gtri(src, dst, old_int_run, old_dec_run, new_int_run, new_dec_run):
    """Copy ``src`` docx to ``dst`` with the Portfolio Highlights GTRI value runs
    swapped, to simulate a highlights figure that has drifted away from the body."""
    import os
    import shutil

    shutil.copy2(src, dst)
    work = dst.parent / '_unz'
    if work.exists():
        shutil.rmtree(work)
    work.mkdir()
    with zipfile.ZipFile(dst) as z:
        z.extractall(work)
    doc_path = work / 'word' / 'document.xml'
    doc = doc_path.read_text('utf-8')
    i = doc.find(old_int_run)
    doc = doc[:i] + new_int_run + doc[i + len(old_int_run):]
    j = doc.find(old_dec_run, i)
    doc = doc[:j] + new_dec_run + doc[j + len(old_dec_run):]
    doc_path.write_text(doc, 'utf-8')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for root, _, files in os.walk(work):
            for f in files:
                fp = Path(root) / f
                zo.write(fp, fp.relative_to(work))


def test_highlights_gtri_updates_when_divergent_from_body(tmp_path):
    """The highlights-box GTRI must update even when it has drifted to a value far
    from the body GTRI (outside the value-match tolerance) — the reported defect."""
    base = Path('ref-files/Quarterly QSP - Q1 2026 - Draft.docx')
    tmpl = tmp_path / 'divergent.docx'
    # Body GTRI stays 3,332; highlights GTRI -> 3,273.6 (stale, >0.25 away).
    _set_highlights_gtri(
        base, tmpl, '>3,331.</w:t>', '>9</w:t>', '>3,273.</w:t>', '>6</w:t>'
    )

    ma_values = {
        'gtri': 3400.0, 'gross_rental': 3120.0, 'vacancy_pct': 6.2,
        'vacancy_amount': 210.0, 'maintenance': 260.0,
        'unit_sales_proceeds': 203.5, 'gtri_ltm': 13500.0,
        'gross_rental_ltm': 12800.0, 'maintenance_ltm': 1000.0,
    }
    rv = build_report_values(
        ma_values=ma_values, rent_roll_k=13317.9, rent_roll_units=400,
        units_sold_quarter=2, unit_sales_proceeds=203.5,
        word_template_path=tmpl, report_date=datetime(2026, 4, 15),
    )
    out = tmp_path / 'out.docx'
    WordTemplateUpdater(str(tmpl), str(out)).update_with_python_docx(
        rv, 'Q4 2025', 'Q1 2026'
    )
    text = _visible_text(out)

    assert '3,273.6' not in text, 'stale highlights GTRI survived'
    m = re.search(r'Gross Theoretical rental income:\s*€?\s?([\d,\.]+)', text)
    assert m and m.group(1).replace(',', '').startswith('3400'), m.group(1) if m else 'no match'
    assert not re.search(r'\d\.\d\.\d', text)
