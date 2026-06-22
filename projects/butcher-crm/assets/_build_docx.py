# -*- coding: utf-8 -*-
"""Markdown → Word(.docx). 이미지·각주·페이지나눔 포함."""
import re, sys, pathlib, pypandoc

BASE = pathlib.Path("projects/butcher-crm")

# <!--PAGE--> → Word 페이지 나눔(raw openxml)
PAGEBREAK = "\n\n```{=openxml}\n<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>\n```\n\n"

def convert(src, out):
    md = pathlib.Path(src).read_text(encoding="utf-8")
    md = md.replace("<!--PAGE-->", PAGEBREAK)
    # 나머지 HTML 주석 제거 (단, openxml 블록은 유지)
    md = re.sub(r"<!--(?!PAGE).*?-->", "", md, flags=re.DOTALL)
    pypandoc.convert_text(
        md, "docx", format="markdown+pipe_tables+footnotes+raw_attribute",
        outputfile=str(out),
        extra_args=[f"--resource-path={BASE}", "--standalone"],
    )
    print("saved", out)

JOBS = {
    "submit": (BASE/"예비창업패키지_PSST_제출배치판_v1.0.md",
               BASE/"예비창업패키지_PSST_제출배치판_v1.0.docx"),
    "master": (BASE/"예비창업패키지_PSST_사업계획서_v0.1.md",
               BASE/"예비창업패키지_PSST_사업계획서_마스터.docx"),
}
for k in (sys.argv[1:] or list(JOBS)):
    convert(*JOBS[k])
