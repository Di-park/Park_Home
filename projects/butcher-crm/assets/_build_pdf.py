# -*- coding: utf-8 -*-
"""Markdown → PDF (미리보기용). NanumGothic 임베드, 이미지 포함."""
import re, sys, markdown, pathlib
from weasyprint import HTML

BASE = pathlib.Path("projects/butcher-crm")
FONT_DIR = "/usr/local/lib/python3.11/dist-packages/koreanize_matplotlib/fonts"

CSS = f"""
@font-face {{ font-family:'Nanum'; src:url('file://{FONT_DIR}/NanumGothic.ttf'); font-weight:normal; }}
@font-face {{ font-family:'Nanum'; src:url('file://{FONT_DIR}/NanumGothicBold.ttf'); font-weight:bold; }}
@page {{ size:A4; margin:1.5cm 1.4cm; }}
* {{ font-family:'Nanum', sans-serif; }}
body {{ font-size:9.5pt; line-height:1.65; color:#1A2A3A; }}
h1 {{ font-size:15pt; border-bottom:2.5px solid #2A5C93; padding-bottom:4px; margin:16px 0 10px; color:#1d3a5f; }}
h2 {{ font-size:12.5pt; margin:14px 0 6px; padding-left:7px; border-left:4px solid #2A5C93; color:#234; }}
h3 {{ font-size:10.8pt; margin:10px 0 5px; }}
h4 {{ font-size:10.2pt; margin:11px 0 5px; color:#2A5C93; }}
p {{ margin:5px 0; }}
ul, ol {{ margin:5px 0 8px; padding-left:20px; }}
li {{ margin:3px 0; line-height:1.55; }}
table {{ border-collapse:collapse; width:100%; margin:8px 0 10px; font-size:8.6pt; page-break-inside:avoid; }}
th, td {{ border:1px solid #B8C2CC; padding:4px 6px; text-align:left; vertical-align:top; line-height:1.45; }}
th {{ background:#EAF0F7; font-weight:bold; }}
img {{ max-width:92%; margin:8px auto; display:block; page-break-inside:avoid; }}
blockquote {{ margin:7px 0; padding:5px 12px; border-left:3px solid #E8862E; color:#5a4632;
            font-size:8.8pt; background:#FFF8F0; border-radius:2px; }}
code, pre {{ font-family:'Nanum'; font-size:8pt; background:#F4F6F8; white-space:pre-wrap; }}
hr {{ border:none; border-top:1px solid #D8DEE5; margin:10px 0; }}
.pagebreak {{ page-break-before:always; }}
.footnote {{ font-size:7.8pt; color:#54606C; }}
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
