# -*- coding: utf-8 -*-
"""Markdown → PDF (미리보기용). NanumGothic 임베드, 이미지 포함."""
import re, sys, markdown, pathlib
from weasyprint import HTML

BASE = pathlib.Path("projects/butcher-crm")
FONT_DIR = "/usr/local/lib/python3.11/dist-packages/koreanize_matplotlib/fonts"

CSS = f"""
@font-face {{ font-family:'Nanum'; src:url('file://{FONT_DIR}/NanumGothic.ttf'); font-weight:normal; }}
@font-face {{ font-family:'Nanum'; src:url('file://{FONT_DIR}/NanumGothicBold.ttf'); font-weight:bold; }}
@page {{ size:A4; margin:1.4cm 1.3cm; }}
* {{ font-family:'Nanum', sans-serif; }}
body {{ font-size:9.3pt; line-height:1.5; color:#1A2A3A; }}
h1 {{ font-size:15pt; border-bottom:2px solid #2A5C93; padding-bottom:3px; margin:14px 0 8px; color:#1d3a5f; }}
h2 {{ font-size:12.5pt; margin:12px 0 6px; color:#234; }}
h3 {{ font-size:10.8pt; margin:9px 0 4px; }}
h4 {{ font-size:9.8pt; margin:7px 0 3px; }}
p {{ margin:4px 0; }}
table {{ border-collapse:collapse; width:100%; margin:6px 0; font-size:8.4pt; }}
th, td {{ border:1px solid #B8C2CC; padding:3px 5px; text-align:left; vertical-align:top; }}
th {{ background:#EAF0F7; font-weight:bold; }}
img {{ max-width:100%; margin:6px auto; display:block; }}
blockquote {{ margin:5px 0; padding:3px 10px; border-left:3px solid #9DB8DC; color:#42505E; font-size:8.4pt; background:#F6F8FB; }}
code, pre {{ font-family:'Nanum'; font-size:8pt; background:#F4F6F8; white-space:pre-wrap; }}
hr {{ border:none; border-top:1px solid #D8DEE5; margin:8px 0; }}
.pagebreak {{ page-break-before:always; }}
"""

def convert(src, out):
    md = pathlib.Path(src).read_text(encoding="utf-8")
    md = md.replace("<!--PAGE-->", '\n\n<div class="pagebreak"></div>\n\n')
    md = re.sub(r"<!--.*?-->", "", md, flags=re.DOTALL)
    body = markdown.markdown(md, extensions=["tables", "fenced_code", "footnotes"])
    html = f"<html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>"
    HTML(string=html, base_url=str(BASE)).write_pdf(str(out))
    print("saved", out)

JOBS = {
    "submit": (BASE/"예비창업패키지_PSST_제출배치판_v1.0.md",
               BASE/"예비창업패키지_PSST_제출배치판_v1.0.pdf"),
    "master": (BASE/"예비창업패키지_PSST_사업계획서_v0.1.md",
               BASE/"예비창업패키지_PSST_사업계획서_마스터.pdf"),
}
sel = sys.argv[1:] or list(JOBS)
for k in sel:
    convert(*JOBS[k])
