# -*- coding: utf-8 -*-
"""소상공인 고객관리 AI 에이전트 — 시스템 아키텍처 도식 생성"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib  # 나눔고딕 자동 적용
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

# 색상
C_NORMAL = "#E8EEF7"
C_NORMAL_EDGE = "#4A6FA5"
C_CORE = "#FFE8D6"
C_CORE_EDGE = "#E8862E"
C_OUT = "#E3F1E6"
C_OUT_EDGE = "#3C8C57"
C_TEXT = "#1A2A3A"

fig, ax = plt.subplots(figsize=(14, 6.4))
ax.set_xlim(0, 14)
ax.set_ylim(0, 6.4)
ax.axis("off")

def box(cx, cy, w, h, title, body, face, edge, title_size=13, body_size=10.5):
    ax.add_patch(FancyBboxPatch((cx - w/2, cy - h/2), w, h,
                 boxstyle="round,pad=0.04,rounding_size=0.18",
                 linewidth=2, facecolor=face, edgecolor=edge, zorder=3))
    if body:
        ax.text(cx, cy + h/2 - 0.34, title, ha="center", va="center",
                fontsize=title_size, fontweight="bold", color=C_TEXT, zorder=4)
        ax.text(cx, cy - 0.12, body, ha="center", va="center",
                fontsize=body_size, color=C_TEXT, zorder=4, linespacing=1.5)
    else:
        ax.text(cx, cy, title, ha="center", va="center",
                fontsize=title_size, fontweight="bold", color=C_TEXT, zorder=4)

def arrow(x1, y1, x2, y2, color="#33475B"):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2),
                 arrowstyle="-|>", mutation_scale=22, linewidth=2.4,
                 color=color, zorder=2, shrinkA=2, shrinkB=2))

# 상단 행 (좌→우)
y_top = 4.5
box(1.7, y_top, 2.5, 1.25, "매장 POS", "거래 데이터", C_NORMAL, C_NORMAL_EDGE)
box(4.9, y_top, 2.4, 1.25, "수집 · 적재", "", C_NORMAL, C_NORMAL_EDGE)
box(8.2, y_top, 2.9, 1.7, "단골 분석 엔진",
    "· 5단계 세그먼트 분류\n· 이탈 예측\n· 챙길 단골 자동 선정", C_CORE, C_CORE_EDGE)
box(11.7, y_top, 2.9, 1.7, "LLM 메시지 생성",
    "+ 안전 가드레일\n(빈도·시간·금칙)", C_CORE, C_CORE_EDGE)

# 하단 행 (우→좌)
y_bot = 1.5
box(11.7, y_bot, 2.6, 1.25, "발송", "알림톡 · 문자", C_OUT, C_OUT_EDGE)
box(8.2, y_bot, 2.6, 1.25, "재방문·매출 추적", "", C_OUT, C_OUT_EDGE)
box(4.9, y_bot, 2.6, 1.45, "주간 리포트",
    "사장님 주 5초 확인", C_OUT, C_OUT_EDGE)

# 화살표
arrow(2.95, y_top, 3.70, y_top)
arrow(6.10, y_top, 6.75, y_top)
arrow(9.65, y_top, 10.25, y_top)
arrow(11.7, y_top - 0.85, 11.7, y_bot + 0.63)   # 4 ↓ 5
arrow(10.40, y_bot, 9.50, y_bot)                 # 5 → 6
arrow(6.90, y_bot, 6.20, y_bot)                  # 6 → 7

# 핵심 엔진 그룹 점선 박스
ax.add_patch(FancyBboxPatch((6.55, 3.42), 6.75, 2.18,
             boxstyle="round,pad=0.02,rounding_size=0.10",
             linewidth=2, linestyle=(0, (6, 4)), facecolor="none",
             edgecolor=C_CORE_EDGE, zorder=1))
ax.text(9.92, 5.78, "AI 에이전트  ·  Zero-Task 자동화 (사장님 행동 ≈ 0)",
        ha="center", va="center", fontsize=12, fontweight="bold", color=C_CORE_EDGE)

# 제목
ax.text(0.2, 6.15, "[그림 2] 소상공인 사업장 특화 고객관리 AI 에이전트 — 시스템 아키텍처",
        ha="left", va="center", fontsize=13.5, fontweight="bold", color=C_TEXT)

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/architecture_diagram.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/architecture_diagram.svg", bbox_inches="tight")
print("saved PNG + SVG")
