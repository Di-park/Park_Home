# -*- coding: utf-8 -*-
"""검증 게이트 (Stage 1→2→3) 도식 ([그림 5])"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

C_TEXT = "#1A2A3A"
STAGE_FACE = ["#E8EEF7", "#CFE0F2", "#9CC2E8"]
STAGE_EDGE = ["#5B8FC9", "#3F77B5", "#2A5C93"]
C_FAIL = "#FBE3E0"; E_FAIL = "#C0473B"
C_GOAL = "#E3F1E6"; E_GOAL = "#3C8C57"

fig, ax = plt.subplots(figsize=(14, 6.2))
ax.set_xlim(0, 14); ax.set_ylim(0, 6.2); ax.axis("off")

def box(cx, cy, w, h, title, sub, body, face, edge, tsize=13):
    ax.add_patch(FancyBboxPatch((cx-w/2, cy-h/2), w, h,
                 boxstyle="round,pad=0.04,rounding_size=0.16",
                 linewidth=2, facecolor=face, edgecolor=edge, zorder=3))
    ax.text(cx, cy+h/2-0.34, title, ha="center", va="center",
            fontsize=tsize, fontweight="bold", color=C_TEXT, zorder=4)
    if sub:
        ax.text(cx, cy+h/2-0.72, sub, ha="center", va="center",
                fontsize=10, color=edge, fontweight="bold", zorder=4)
    if body:
        ax.text(cx, cy-0.28, body, ha="center", va="center",
                fontsize=9.8, color=C_TEXT, zorder=4, linespacing=1.5)

def arrow(x1, y1, x2, y2, color, lw=2.4, style="-|>", ls="-"):
    ax.add_patch(FancyArrowPatch((x1,y1),(x2,y2), arrowstyle=style,
                 mutation_scale=20, linewidth=lw, color=color, zorder=2,
                 linestyle=ls, shrinkA=1, shrinkB=1))

y = 4.4
cxs = [2.4, 6.4, 10.4]
stages = [
    ("Stage 1  가설검증", "+3~6개월",
     "매출 +10%↑\n알림→재방문 8%↑\n이탈임박군 잔존 +20%p"),
    ("Stage 2  PMF", "+6~8개월",
     "30매장 70%↑ 효과\n유료 전환 60%↑\n발송량 분포 안정"),
    ("Stage 3  공헌이익", "후속(~18m)",
     "ARPU 8~12만원\n월 이탈 5%↓\nLTV·CAC 5배+"),
]
for i,(t,s,b) in enumerate(stages):
    box(cxs[i], y, 3.4, 2.0, t, s, b, STAGE_FACE[i], STAGE_EDGE[i])

# 통과 화살표 + 게이트 라벨
for i in range(2):
    arrow(cxs[i]+1.75, y, cxs[i+1]-1.75, y, "#2E7D4F")
    ax.text((cxs[i]+cxs[i+1])/2, y+0.32, "통과", ha="center", va="center",
            fontsize=10.5, fontweight="bold", color="#2E7D4F")
# Stage3 → 후속투자
arrow(cxs[2]+1.75, y, 12.5, y, "#2E7D4F")
box(13.15, y, 1.9, 2.0, "후속\n투자", "", "Pre-Seed\n· R&D", C_GOAL, E_GOAL, tsize=12)

# 실패 분기 (아래로 빨강)
fails = ["실패 → 제품 피벗", "실패 → 가격·세그먼트 재설계", "실패 → 브릿지·피벗"]
for i in range(3):
    arrow(cxs[i], y-1.05, cxs[i], 1.95, E_FAIL, lw=2, ls=(0,(5,3)))
    ax.add_patch(FancyBboxPatch((cxs[i]-1.55, 1.0), 3.1, 0.85,
                 boxstyle="round,pad=0.03,rounding_size=0.12",
                 linewidth=1.6, facecolor=C_FAIL, edgecolor=E_FAIL, zorder=3))
    ax.text(cxs[i], 1.42, fails[i], ha="center", va="center",
            fontsize=10, fontweight="bold", color=E_FAIL, zorder=4)

ax.text(0.1, 5.95, "[그림 5] 검증 게이트 — Stage 1 → 2 → 3 (통과/실패 의사결정)",
        ha="left", va="center", fontsize=14, fontweight="bold", color=C_TEXT)
ax.text(0.1, 0.45, "※ 본 사업비는 '제품을 만드는' 돈이 아니라 각 Stage의 가설에 '측정된 답'을 확보하는 돈",
        ha="left", va="center", fontsize=10, color="#42505E")

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/stage_gate.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/stage_gate.svg", bbox_inches="tight")
print("saved stage gate")
