# -*- coding: utf-8 -*-
"""시장 규모 TAM-SAM-SOM 동심원 도식 ([그림 4])"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch

C_TEXT = "#1A2A3A"
C_TAM = "#DCE7F5"; E_TAM = "#9DB8DC"
C_SAM = "#9CC2E8"; E_SAM = "#5B8FC9"
C_SOM = "#E8862E"; E_SOM = "#C26A18"

fig, ax = plt.subplots(figsize=(12.5, 7))
ax.set_xlim(0, 12.5); ax.set_ylim(0, 7); ax.axis("off")
ax.set_aspect("equal")

cx = 3.7
# 바닥 정렬 동심원 (작을수록 아래)
r_tam, r_sam, r_som = 3.25, 2.25, 1.25
ax.add_patch(Circle((cx, r_tam), r_tam, facecolor=C_TAM, edgecolor=E_TAM, lw=2, zorder=1))
ax.add_patch(Circle((cx, r_sam), r_sam, facecolor=C_SAM, edgecolor=E_SAM, lw=2, zorder=2))
ax.add_patch(Circle((cx, r_som), r_som, facecolor=C_SOM, edgecolor=E_SOM, lw=2, zorder=3))

# 원 안 라벨
ax.text(cx, r_tam*2 - 0.45, "TAM", ha="center", va="center",
        fontsize=15, fontweight="bold", color="#3A5A86")
ax.text(cx, r_sam*2 - 0.40, "SAM", ha="center", va="center",
        fontsize=14, fontweight="bold", color="#274C77")
ax.text(cx, r_som, "SOM", ha="center", va="center",
        fontsize=13, fontweight="bold", color="white")

# 우측 설명 (스와치 + 정의 + 수치)
def legend(y, color, edge, tag, defn, fig_val):
    ax.add_patch(plt.Rectangle((7.7, y - 0.22), 0.44, 0.44, facecolor=color,
                 edgecolor=edge, lw=1.6, zorder=4))
    ax.text(8.35, y + 0.16, tag, ha="left", va="center", fontsize=13.5,
            fontweight="bold", color=C_TEXT)
    ax.text(8.35, y - 0.18, defn, ha="left", va="center", fontsize=10.8, color="#42505E")
    ax.text(8.35, y - 0.56, fig_val, ha="left", va="center", fontsize=12,
            fontweight="bold", color=edge)

legend(5.55, C_TAM, "#3A5A86", "TAM  재방문 기반 B2C 소상공인",
       "음식점·소매·미용·서비스 전반", "약 300만 업체 · 7조원")
legend(3.55, C_SAM, "#274C77", "SAM  식품 소매·외식",
       "정육 도메인 엔진 재사용 가능권", "약 90만 개 · 1조원")
legend(1.55, C_SOM, "#C26A18", "SOM  정육점 (첫 진입)",
       "단골 의존도·데이터 깊이 최적", "약 57,000개소 · 2,200억원")

# 진입→확장 화살표 (좌측, SOM에서 시작해 바깥으로)
ax.annotate("", xy=(0.55, 6.0), xytext=(0.55, 1.0),
            arrowprops=dict(arrowstyle="-|>", color="#C0473B", lw=2.6))
ax.text(0.95, 3.5, "정육(SOM)에서 시작 → 인접 시장 확장", ha="center", va="center",
        fontsize=11, fontweight="bold", color="#C0473B", rotation=90)

ax.text(0.1, 6.75, "[그림 4] 시장 규모 (TAM–SAM–SOM)", ha="left", va="center",
        fontsize=14, fontweight="bold", color=C_TEXT)
ax.text(7.7, 0.5, "※ 시장금액은 매장 수 × ARPU 가정 기반 자체 추정",
        ha="left", va="center", fontsize=9.5, color="#8A97A4")

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/market_tam_sam_som.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/market_tam_sam_som.svg", bbox_inches="tight")
print("saved market diagram")
