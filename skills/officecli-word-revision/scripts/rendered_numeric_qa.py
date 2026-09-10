"""Strict decimal-token inventory for an explicitly bounded rendered table.

Input must be Word PDF text lines, not DOCX or a pre-render preview. This checks
decimal tokens/counts (including signs), not causal validity or all layout rules.
Never join lines to make a split token pass. The adapter owns page/table bounds.
"""
from collections import Counter
from decimal import Decimal
import re
from zipfile import ZipFile
from lxml import etree

PATTERN=re.compile(r'(?<![\w.])[-+−]?\d*\.\d+(?![\w.])')
def tokens(text):
 return Counter(str(Decimal(x.replace('−','-')).normalize()) for x in PATTERN.findall(text))
def check_decimal_inventory(expected_cells,rendered_lines):
 expected=sum((tokens(s) for s in expected_cells),Counter())
 observed=sum((tokens(s) for s in rendered_lines),Counter())
 return {'all_pass':expected==observed and bool(expected),'expected_tokens':sum(expected.values()),'observed_tokens':sum(observed.values()),'missing':dict(expected-observed),'unexpected':dict(observed-expected),'scope':'decimal tokens in explicitly bounded Word-rendered table; no cross-line joining'}
def story_text(path):
 ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 with ZipFile(path) as z:
  names=[n for n in z.namelist() if n in ('word/document.xml','word/footnotes.xml','word/endnotes.xml') or re.fullmatch(r'word/(header|footer)\d+\.xml',n)]
  return {n:[''.join(p.xpath('.//w:t/text()',namespaces=ns))for p in etree.fromstring(z.read(n)).xpath('.//w:p',namespaces=ns)]for n in names}
