from pdfminer.high_level import extract_text
from pathlib import Path
p = Path('excel/楽天関数/ms2rss_function.pdf')
print(extract_text(str(p)))
