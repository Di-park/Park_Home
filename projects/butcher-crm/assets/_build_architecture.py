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
C_CORE = "#FFE8D6"          # AI (주황)
C_CORE_EDGE = "#E8862E"
C_HUMAN = "#FFF6CC"         # 사장님 승인 (노랑)
C_HUMAN_EDGE = "#C9A227"
C_OUT = "#E3F1E6"           # 결과 (초록)
C_OUT_EDGE = "#3C8C57"
C_TEXT = "#1A2A3A"

fig, ax = plt.subplots(figsize=(14, 6.6))
ax.set_xlim(0, 14)
ax.set_ylim(0, 6.6)
ax.axis("off")

def box(cx, cy, w, h, title, body, face, edge, title_size=13, body_size=10.3, badge=None):
    ax.add_patch(FancyBboxPatch((cx - w/2, cy - h/2), w, h,
                 boxstyle="round,pad=0.04,rounding_size=0.18",
                 linewidth=2, facecolor=face, edgecolor=edge, zorder=3))
    if body:
        ax.text(cx, cy + h/2 - 0.36, title, ha="center", va="center",
                fontsize=title_size, fontweight="bold", color=C_TEXT, zorder=4)
        ax.text(cx, cy - 0.16, body, ha="center", va="center",
                fontsize=body_size, color=C_TEXT, zorder=4, linespacing=1.5)
    else:
        ax.text(cx, cy, title, ha="center", va="center",
                fontsize=title_size, fontweight="bold", color=C_TEXT, zorder=4)
    if badge:
        ax.text(cx - w/2 + 0.05, cy + h/2 + 0.02, badge, ha="center", va="center",
                fontsize=12.5, fontweight="bold", color="white", zorder=6,
                bbox=dict(boxstyle="circle,pad=0.22", facecolor=edge, edgecolor="none"))

def arrow(x1, y1, x2, y2, color="#33475B"):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2),
                 arrowstyle="-|>", mutation_scale=22, linewidth=2.4,
                 color=color, zorder=2, shrinkA=2, shrinkB=2))

# 상단 행 (좌→우)
y_top = 4.5
box(1.7, y_top, 2.5, 1.25, "매장 POS", "거래 데이터", C_NORMAL, C_NORMAL_EDGE)
box(4.85, y_top, 2.3, 1.25, "수집 · 적재", "", C_NORMAL, C_NORMAL_EDGE)
box(8.15, y_top, 3.0, 1.75, "AI 진단",
    "· 5단계 세그먼트·이탈 예측\n· 챙길 단골 선정\n· 보낼 메시지 제안", C_CORE, C_CORE_EDGE, badge="①")
box(11.75, y_top, 2.8, 1.75, "사장님 OK",
    "AI 제안 확인·승인\n(탭 1번)", C_HUMAN, C_HUMAN_EDGE, badge="②")

# 하단 행 (우→좌)
y_bot = 1.45
box(11.75, y_bot, 3.0, 1.75, "AI 실행",
    "승인 고객에게\nLLM 메시지 발송\n(알림톡·문자)", C_CORE, C_CORE_EDGE, badge="③")
box(8.15, y_bot, 2.6, 1.25, "재방문·매출 추적", "", C_OUT, C_OUT_EDGE)
box(4.85, y_bot, 2.6, 1.5, "주간 리포트",
    "사장님 주 5초 확인", C_OUT, C_OUT_EDGE)

# 화살표
arrow(2.95, y_top, 3.70, y_top)
arrow(6.00, y_top, 6.65, y_top)
arrow(9.65, y_top, 10.35, y_top)
arrow(11.75, y_top - 0.88, 11.75, y_bot + 0.88)   # 사장님 OK ↓ AI 실행
arrow(10.25, y_bot, 9.45, y_bot)                  # AI 실행 → 추적
arrow(6.85, y_bot, 6.15, y_bot)                   # 추적 → 리포트

# AI 영역 점선 박스 2개 (진단 / 실행)
for (bx, by, bw, bh) in [(8.15, y_top, 3.25, 2.05), (11.75, y_bot, 3.25, 2.05)]:
    ax.add_patch(FancyBboxPatch((bx - bw/2, by - bh/2), bw, bh,
                 boxstyle="round,pad=0.02,rounding_size=0.10",
                 linewidth=1.8, linestyle=(0, (5, 4)), facecolor="none",
                 edgecolor=C_CORE_EDGE, zorder=1))

# 제목 + 부제
ax.text(0.2, 6.35, "[그림 2] 소상공인 사업장 특화 고객관리 AI 에이전트 — 시스템 아키텍처",
        ha="left", va="center", fontsize=13.5, fontweight="bold", color=C_TEXT)
ax.text(0.2, 5.85, "AI가 ① 진단(제안) → 사장님이 ② OK → AI가 ③ 실행(발송).  "
        "진단·작성·발송은 AI, 사장님은 'OK' 한 번 (Zero-Task)",
        ha="left", va="center", fontsize=11.5, color=C_CORE_EDGE, fontweight="bold")

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/architecture_diagram.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/architecture_diagram.svg", bbox_inches="tight")
print("saved PNG + SVG")
