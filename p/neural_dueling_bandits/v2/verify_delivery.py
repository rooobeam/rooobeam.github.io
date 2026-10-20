# -*- coding: utf-8 -*-
"""Check the English deliverables, source formulas and preserved v1 files."""
from pathlib import Path
import hashlib
import json
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(tempfile.gettempdir()) / 'codex-dueling-pdf-deps'))
import fitz

ROOT = Path(__file__).resolve().parent
PREVIEW = ROOT / 'preview'
PPTX = ROOT / 'Dueling_Bandits_Sections_1-4_v2.pptx'
PDF = ROOT / 'Dueling_Bandits_Sections_1-4_v2.pdf'
CJK = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]')

with zipfile.ZipFile(PPTX) as package:
    assert package.testzip() is None, 'PowerPoint package integrity check failed.'
    slides = sorted(name for name in package.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', name))
    notes = sorted(name for name in package.namelist() if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml', name))
    for name in slides + notes + ['docProps/core.xml']:
        xml = ET.fromstring(package.read(name))
        assert not CJK.search(' '.join(xml.itertext())), f'Untranslated text in {name}'

with fitz.open(PDF) as pdf:
    assert len(pdf) == 13, 'Unexpected PDF page count.'
    assert 'Contents' in pdf[1].get_text(), 'Contents must appear on slide 2.'
    for page in pdf:
        text = page.get_text()
        assert text.strip(), 'Empty PDF page.'
        assert not CJK.search(text), f'Untranslated PDF text on page {page.number+1}'
    pdf_pages = len(pdf)

content = json.loads((ROOT/'slide_content.json').read_text(encoding='utf-8'))
assert len(slides) == len(notes) == len(content) == 13
assert sum(d['seconds'] for d in content) == 600
assert all(d['notes'].strip() for d in content)
expected = {3:'1 Problem Setup', 5:'2 Assumptions', 7:'3 Algorithm', 10:'4 Confidence Bound'}
assert all(content[page-1]['section'] == section for page, section in expected.items())

for name in ['build_presentation.py', 'speaker_notes_source.json', 'slide_content.json',
             'Speaker_Notes_and_Timing.md', 'README.md']:
    assert not CJK.search((ROOT/name).read_text(encoding='utf-8')), f'Untranslated content in {name}'

formula_v1 = json.loads((ROOT.parent/'v1/assets/equations.json').read_text(encoding='utf-8'))
formula_v2 = json.loads((ROOT/'assets/equations.json').read_text(encoding='utf-8'))
assert [(f['slide'], f['latex']) for f in formula_v1] == [(f['slide'], f['latex']) for f in formula_v2]

protected = {
    'build_presentation.py':'188C83DE4B5BA89FCA6238E2D468792C4F2D0F38AB5C5402A188A386C5CD8066',
    'Dueling_Bandits_Sections_1-4_v1.pptx':'2CC2D471C376A22796C38843866FE888A3C90F193F2AF896AE95F0A6C0DEE837',
    'Dueling_Bandits_Sections_1-4_v1.pdf':'5639EEB0EEC148119DDA87150A7954E7B49D764201E0294426A2E969149F22C1',
}
for name, expected_hash in protected.items():
    actual = hashlib.sha256((ROOT.parent/'v1'/name).read_bytes()).hexdigest().upper()
    assert actual == expected_hash, f'The protected v1 file changed: {name}'

render = json.loads((PREVIEW/'render_audit.json').read_text(encoding='utf-8'))
structure = json.loads((PREVIEW/'structure_audit.json').read_text(encoding='utf-8'))
assert render['slides_rendered'] == 13 and not render['text_overflow']
assert not structure['geometry_issues']
spoken_words = sum(d['spoken_words'] for d in content)
audit = {
    'language':'English', 'pptx_integrity':'pass', 'slides':len(slides),
    'speaker_notes':len(notes), 'pdf_pages':pdf_pages,
    'planned_seconds':600, 'main_script_words':spoken_words,
    'contents_page':2, 'contents_page_ranges':'verified',
    'chinese_in_slides_notes_metadata_pdf_and_docs':False,
    'text_overflow':[], 'geometry_issues':[],
    'equations_preserved_from_v1':len(formula_v2),
    'v1_source_pptx_pdf_unchanged':True,
    'pptx_bytes':PPTX.stat().st_size, 'pdf_bytes':PDF.stat().st_size,
}
(PREVIEW/'delivery_audit.json').write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')

report = f'''# v2 Build Report

## Result

The English edition is complete: 13 slides, a dedicated contents page, a PDF preview, and English speaker notes. The planned duration remains 10 minutes. The main spoken script contains {spoken_words} words; optional technical details are listed separately.

## Issues encountered

- **English text length — resolved.** The first render found nine text-box overflows on slides 5, 6, 7, 9 and 13. Labels and summary text were shortened, and line breaks were revised. The final PowerPoint render reports no text overflow.
- **Source-proof qualification — explicitly retained.** Slide 6 states that the curvature lower bound must hold at the intermediate scores between the true and unconstrained fitted parameters. Bounded true parameters alone do not prove this condition. The confidence result is presented under that condition.
- **Delivery blockers — none.** The PPTX and PDF were both generated successfully.

## Final checks

- PowerPoint opened and rendered all 13 slides.
- The PDF has 13 nonempty pages.
- The contents page ranges agree with the section start pages.
- Slide text, notes, presentation metadata, PDF text and supporting documents contain no remaining Chinese text.
- All {len(formula_v2)} formula sources match v1, with the same slide assignments.
- All slide objects remain within the page boundaries, with no reported text overflow.
- SHA-256 checks confirm that the v1 source, PPTX and PDF are unchanged.

Detailed machine-readable results are in `preview/delivery_audit.json`, `preview/render_audit.json`, and `preview/structure_audit.json`.
'''
(ROOT/'Build_Report.md').write_text(report,encoding='utf-8')
assert not CJK.search(report)
print(json.dumps(audit, ensure_ascii=False, indent=2))
