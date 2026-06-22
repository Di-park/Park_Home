# -*- coding: utf-8 -*-
"""사업 추진 일정 간트 (협약 약 8개월) ([그림 6])"""
import matplotlib
matplotlib.use("Agg")
import koreanize_matplotlib
import matplotlib.pyplot as plt

C_TEXT = "#1A2A3A"

# (작업명, 시작월, 종료월, 색)
tasks = [
    ("효과 측정 인프라 구축",        1, 4, "#5B8FC9"),
    ("파일럿 효과 데이터 정리",      3, 6, "#5B8FC9"),
    ("협력 5 + 베타 확장(N=30)",    3, 8, "#E8862E"),
    ("AI 추천·발송 고도화",          3, 8, "#E8862E"),
    ("유료 전환·PMF 1차 검증",       5, 8, "#3C8C57"),
    ("성과 정리·후속 사업화 준비",   7, 8, "#7B5EA7"),
]

fig, ax = plt.subplots(figsize=(13, 5.4))
n = len(tasks)
for i, (name, s, e, c) in enumerate(tasks):
    y = n - i
    ax.barh(y, e - s + 1, left=s - 0.5, height=0.55, color=c,
            edgecolor="white", zorder=3)
    ax.text(s - 0.5 + (e - s + 1)/2, y, f"M{s}~M{e}", ha="center", va="center",
            fontsize=9.5, color="white", fontweight="bold", zorder=4)
    ax.text(0.4, y, name, ha="left", va="center", fontsize=11, color=C_TEXT)

# 마일스톤 마커
ms = {4: "S1 게이트", 6: "베타 N=30", 8: "S2·후속 준비"}
for m, lab in ms.items():
    ax.axvline(m + 0.5, color="#C0473B", lw=1.4, ls=(0, (4, 3)), zorder=2)
    ax.text(m + 0.5, n + 0.75, lab, ha="center", va="bottom",
            fontsize=9.5, color="#C0473B", fontweight="bold")

ax.set_xlim(0, 9)
ax.set_ylim(0.3, n + 1.3)
ax.set_xticks(range(1, 9))
ax.set_xticklabels([f"M{i}" for i in range(1, 9)], fontsize=10.5)
ax.set_yticks([])
ax.set_xlabel("협약기간 (약 8개월)", fontsize=11, color=C_TEXT)
for spdir in ["top", "right", "left"]:
    ax.spines[spdir].set_visible(False)
ax.grid(axis="x", color="#E5E5E5", zorder=0)
ax.set_axisbelow(True)

ax.set_title("[그림 6] 사업 추진 일정 (협약 약 8개월) — 본 사업 목표: Stage 1 완주 + Stage 2 진입",
             fontsize=13.5, fontweight="bold", color=C_TEXT, loc="left", pad=24)

plt.tight_layout()
fig.savefig("projects/butcher-crm/assets/gantt.png", dpi=300, bbox_inches="tight")
fig.savefig("projects/butcher-crm/assets/gantt.svg", bbox_inches="tight")
print("saved gantt")
