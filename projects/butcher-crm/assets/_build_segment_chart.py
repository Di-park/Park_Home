# -*- coding: utf-8 -*-
"""파일럿 매장 고객 세그먼트 — 막대그래프 ([그림 3])"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib
import matplotlib.pyplot as plt

segs   = ["완전단골", "단골", "일반활성", "이탈위험", "완전이탈"]
counts = [332, 269, 551, 1088, 2255]
pct    = [7.4, 6.0, 12.3, 24.2, 50.2]
ltv_man = [103.0, 47.8, 28.5, 21.6, 11.9]   # 1인 누적구매(만원)
colors = ["#2E7D4F", "#5BA86F", "#E0B33A", "#E2843A", "#C0473B"]
C_TEXT = "#1A2A3A"

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13.5, 5.6))

# (좌) 고객수
b1 = ax1.bar(segs, counts, color=colors, width=0.66, zorder=3)
for i, r in enumerate(b1):
    ax1.text(r.get_x() + r.get_width()/2, r.get_height() + 30,
             f"{counts[i]:,}명\n({pct[i]}%)", ha="center", va="bottom",
             fontsize=10.5, fontweight="bold", color=C_TEXT)
ax1.set_title("(a) 고객수 — 총 4,495명", fontsize=13, fontweight="bold", pad=14)
ax1.set_ylim(0, 2750)
ax1.grid(axis="y", color="#DDDDDD", zorder=0)
ax1.set_axisbelow(True)
for s in ["top", "right"]:
    ax1.spines[s].set_visible(False)

# 이탈·위험 74.4% 강조 브래킷
ax1.annotate("", xy=(2.62, 2520), xytext=(4.38, 2520),
             arrowprops=dict(arrowstyle="-", color="#C0473B", lw=2))
ax1.text(3.5, 2600, "이탈했거나 이탈 위험 = 74.4% (3,343명)",
         ha="center", va="bottom", fontsize=11, fontweight="bold", color="#C0473B")

# (우) 1인 누적구매 (LTV)
b2 = ax2.bar(segs, ltv_man, color=colors, width=0.66, zorder=3)
for i, r in enumerate(b2):
    ax2.text(r.get_x() + r.get_width()/2, r.get_height() + 1.5,
             f"{ltv_man[i]:.1f}만원", ha="center", va="bottom",
             fontsize=10.5, fontweight="bold", color=C_TEXT)
ax2.set_title("(b) 1인 누적구매 (LTV)", fontsize=13, fontweight="bold", pad=14)
ax2.set_ylim(0, 122)
ax2.grid(axis="y", color="#DDDDDD", zorder=0)
ax2.set_axisbelow(True)
for s in ["top", "right"]:
    ax2.spines[s].set_visible(False)
# 8.7배 격차 표기
ax2.annotate("완전단골 = 완전이탈의 8.7배", xy=(0, 103), xytext=(1.7, 108),
             fontsize=11, fontweight="bold", color="#2E7D4F",
             arrowprops=dict(arrowstyle="->", color="#2E7D4F", lw=1.8))

fig.suptitle("[그림 3] 파일럿 매장 고객 세그먼트 분석 (정육점 1개소, 거래 56,402건)",
             fontsize=14, fontweight="bold", x=0.02, ha="left")
plt.tight_layout(rect=[0, 0, 1, 0.95])
fig.savefig("projects/butcher-crm/assets/segment_chart.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/segment_chart.svg", bbox_inches="tight")
print("saved segment chart")
