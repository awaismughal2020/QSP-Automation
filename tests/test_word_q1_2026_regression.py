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

    # 4. Highlights GTRI matches the real GTRI, not maintenance/proceeds
    m = re.search(r'Gross Theoretical rental income:\s*€?\s?([\d,\.]+)', text)
    assert m and m.group(1).replace(',', '').startswith('3'), m.group(1) if m else 'no match'
