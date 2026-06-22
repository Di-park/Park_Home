# -*- coding: utf-8 -*-
"""제품 화면 목업 — 초안용 빈칸 placeholder ([그림 1])"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

C_TEXT = "#1A2A3A"
C_GRAY = "#9AA7B4"

fig, ax = plt.subplots(figsize=(12, 5.4))
ax.set_xlim(0, 12); ax.set_ylim(0, 5.4); ax.axis("off")

def placeholder(cx, cy, w, h, caption):
    ax.add_patch(FancyBboxPatch((cx - w/2, cy - h/2), w, h,
                 boxstyle="round,pad=0.04,rounding_size=0.12",
                 linewidth=2, linestyle=(0, (6, 5)),
                 facecolor="#F5F7FA", edgecolor=C_GRAY, zorder=2))
    ax.text(cx, cy + 0.35, "[ 제품 화면 — 추후 삽입 ]", ha="center", va="center",
            fontsize=13, fontweight="bold", color=C_GRAY)
    ax.text(cx, cy - 0.45, caption, ha="center", va="center",
            fontsize=11.5, color=C_TEXT)

placeholder(3.1, 2.7, 5.0, 3.6, "① 주간 리포트 화면 예시")
placeholder(8.9, 2.7, 5.0, 3.6, "② 자동 발송 메시지 예시")

ax.text(0.1, 5.15, "[그림 1] 제품 화면 (초안 — 이미지 추후 삽입)",
        ha="left", va="center", fontsize=13.5, fontweight="bold", color=C_TEXT)

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/product_mockup.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/product_mockup.svg", bbox_inches="tight")
print("saved mockup placeholder")
