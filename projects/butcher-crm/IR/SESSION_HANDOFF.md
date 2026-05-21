# 또와 IR 덱 — 세션 핸드오프 (Cold Start Brief)

> 이 문서는 새 Claude Code 세션에서 *또와 IR 덱 작업을 이어갈 때* 필요한 컨텍스트를 압축한 것입니다.
> 새 세션 첫 메시지에 *이 파일을 읽어달라*고 하면 컨텍스트 복원 가능.

---

## 1. 프로젝트 한 줄 정의

**㈜위브원** (대표 안동철) — *정육점 특화 Zero-Task AI 에이전트* **또와**
> 단골부터 마진까지, 알아서 다 챙기는 막내 직원.

- 회사: ㈜위브원 / 서비스명: **또와** (가칭 "단골이" → 확정)
- 팀 5명: 안동철(대표·SaaS), 배현혜(AI·Columbia 박사), SEO CHARLES(제품·KAIST), 이지백(현장·정육 도메인), 박동일(운영)
- 협력사: **㈜더담우** (연 매출 1,000억 정육 유통사, 5매장 영업망 합의)
- 파일럿: 정육점 1매장 (거래 5만 건 / 단골 2,100명) — *매출 효과 미측정*
- 라운드: **Pre-Seed 5억 (희석) + 정부 R&D 7억 (비희석, 디딤돌 2.668억 확보)**

---

## 2. 현재 상태 (최신 산출물)

| 항목 | 경로 |
|------|------|
| 최신 PPT | `/tmp/butcher-build/또와_IR덱_v1.1_AWS.pptx` (15장, AWS 톤, Noto Sans KR) |
| PPT 빌더 | `/tmp/butcher-build/make_pptx_v08.py` (python-pptx, ~1400줄) |
| audit 스크립트 | `/tmp/butcher-build/audit_pptx.py` |
| 수정 노트 | `/tmp/butcher-build/pptx_revisions.md` (슬라이드별 결정사항) |
| 마크다운 IR 덱 | `projects/butcher-crm/IR/또와_IR덱_v0.7.md` ✅ 레포 커밋됨 |
| 사업계획서 | `projects/butcher-crm/단골이_사업계획서_v0.6.md` ⚠️ v0.7 동기화 미완 |
| 디딤돌 R&D 원문 | `research/references/2026-04_디딤돌_연구개발계획서_원문정리.md` (다른 브랜치 origin/claude/ai-crm-rd-plan-yaH7I) |
| 레포 | `di-park/park_home` · 브랜치 `claude/build-ir-deck-ri0ZR` |

---

## 3. 디자인 시스템 (AWS 톤)

```python
# 색상
TOSS_BLUE        = #D15127  # AWS Orange (메인) — 변수명은 레거시
TOSS_BLUE_DARK   = #A8401E  # 진한 오렌지
TOSS_BLUE_LIGHT  = #FAD9CB  # 연한 (강조 배경)
TOSS_BLUE_FAINT  = #FFF7F2  # 매우 옅은 (카드 배경)
TOSS_BLUE_MID    = #E89476  # 중간
AWS_DARK_NAVY    = #232F3E  # 보조 (돈 흐름·강조 텍스트)
# + Gray scale 050~900, GREEN/ORANGE/AMBER/PURPLE/BROWN 보조

# 폰트: Noto Sans KR
# 핵심 함수 _force_font_xml(run) — <a:latin/ea/cs> 3 슬롯 모두 강제 명시

# 슬라이드: 16:9 (13.333" × 7.5"), 흰 배경, 푸터 없음 (페이지 번호만)
```

---

## 4. 15 슬라이드 구성 (현재 v1.1)

| # | 슬라이드 | 핵심 메시지 |
|---|---------|-----------|
| 1 | Cover | 정육점 특화 Zero-Task AI 에이전트 / 단골부터 마진까지 막내 직원 |
| 2 | One-Liner | 첫 AI 직원 / 0분·0분·**하루 5초** 약속 |
| 3 | Problem | 매출 70%가 단골인데 사장님은 모름 / 3 문제 |
| 4 | Solution | 사장님은 눈앞의 손님만 / 효과 3(하루 5초/−50%/+15%) + 기능 3 |
| 5 | Why Now | AI 도메인 특화 시대 / 소상공인 자리 비어있음 |
| 6 | Competition | 수직×Zero-Task 빈자리 / 2×2 맵 + 5플레이어 표 |
| **7** | **Market Insight** | **큰 경쟁자도 못 뚫은 정육점 / "잘 되면 안 바꾼다" / 비용 절감 톤** |
| 8 | Market | TAM 7조 / SAM 1조 / SOM 2,200억 / Capturable 3→10→20% (Square 벤치마크) |
| 9 | BM | 5 참여자 BM Flow (정육 도매상 텍스트 메모로 처리) + 3 카드 |
| 10 | Traction | 5만 거래 / 2,100 단골 / **정직 공개** (효과 미측정) |
| 11 | Validation | 4 Stage: 가설 검증 → PMF → 공헌이익 양수화 → 시장 확장 / Q 내림차순 |
| 12 | Team | 5인 카드 + Founder-Market Fit |
| 13 | Roadmap | 시간선 + 자금 띠 2개 + Series A 진입 조건 |
| 14 | Ask | Pre-Seed 5억 + R&D 7억 / 자금 4 항목 |
| 15 | Vision | Mission-driven — *기술 격차 해소 / 정육점은 그 시작* |

---

## 5. 핵심 결정 사항 (Why)

1. **라운드 사이즈 Pre-Seed 5억** — 시드 12억은 트랙션 갭 (트랙션=N=1 파일럿). Pre-Seed 톤과 매칭.
2. **정부 R&D 7억 별도** — 비희석 자금 (Y Combinator 톤 가산점)
3. **AWS 톤** — 토스 블루(#3182F6)에서 *AWS 오렌지(#D15127)* 로 전면 교체 (사용자 결정)
4. **Solution → Why Now 순서** — 한국말 흐름·청자 호기심 곡선
5. **슬라이드 7 신설** — *큰 경쟁자도 못 뚫은 시장* (defensibility + market insight)
6. **Validation 4단계** — Q1~Q11 (Pre-Seed) + Q12 (Series A 이후)
7. **BM = 단골관리 + 유통 중개 양면 수익** — 정육점에서 유통 매입 + 도매에서 중개 수수료
8. **트랙션 정직 공개** — *매출 효과 미측정* 명시 (founder self-awareness 시그널)
9. **푸터 제거** — 페이지 번호만
10. **비전 = Mission-driven** — *기술 발전의 혜택에서 소외된 소상공인 ...*

---

## 6. 깨짐 방지 5원칙 (PPT 빌더)

1. **폰트**: `<a:latin>` `<a:ea>` `<a:cs>` 3 슬롯 모두 `Noto Sans KR` 명시 (`_force_font_xml`)
2. **이모지 제거**: ✓→[O], 🟡→[!], ✗→[실패] 텍스트 대체. `△`만 안전(Noto Sans CJK KR 포함)
3. **Overflow 방지**: `word_wrap=True` + 박스 크기 폰트×1.4 이상 + `line_spacing` 명시
4. **복잡 도식**: 단순 box + straight connector. L자/곡선/장식 회피. 정육 도매상 같은 *공간 부족 박스*는 텍스트 메모로 대체
5. **표**: native table 아닌 *수동 박스* (corner radius·강조 행 가능) + 컬럼 width 합계 검증

---

## 7. 환경 제약

- LibreOffice headless 변환 *안 됨* (`host_not_allowed`) — PPT 시각 검증은 사용자가 PowerPoint에서 직접
- 외부 호스트 fetch *제한적* (WebFetch 403 흔함)
- `/tmp/butcher-build/` 은 *임시 작업 디렉토리* — 레포에 영구 보관 안 됨
- 사용자 OS: Linux 6.18.5 / Korea timezone / 한국어 입력
- 모델: claude-opus-4-7[1m]
- 폰트 시스템: `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc`

---

## 8. 사용자 작업 스타일

- *한 슬라이드씩 토론* 후 일괄 반영. 트리거: *"이제 피피티 수정해줘"*
- *Solution → Why Now* 같은 *순서 결정*은 사용자가 직감으로 잡음
- *솔직한 비판* 환영 (글로벌 액셀러레이터 톤)
- 슬로건·카피·비전은 *사용자가 직접 제시*하면 그 톤 존중
- 디자인 변경은 *과감하게* (토스→AWS 전환 같은)
- *최대한의 힘* 요청 시 → Plan agent 활용 + 직접 작업 병렬

---

## 9. 다음 작업 후보

1. **사업계획서 v0.7 동기화** — IR 덱 v0.7과 일치 (라운드 5억·BM·비전·슬라이드 7 신설 등)
2. **PPT v1.1 사용자 검증 후 v1.2** — 사용자가 PowerPoint에서 열고 깨진 부분 보고 시
3. **TIPS 신청 자료** — 정부 R&D 7억 확보
4. **파일럿 매장 도입 전/후 매출 측정** (Stage 1 핵심)
5. **외부 모의 발표 2~3회**
6. **앱 스크린샷·데모 영상**

---

## 10. 새 세션 첫 메시지 예시

> "또와 IR 덱 작업을 이어가려고 해. 다음 핸드오프 문서 읽어줘:
> `projects/butcher-crm/IR/SESSION_HANDOFF.md`
>
> 그리고 [구체적 다음 작업]을 진행하자."

또는

> "최신 PPT는 `/tmp/butcher-build/또와_IR덱_v1.1_AWS.pptx`이고, 빌더는 `make_pptx_v08.py`. [수정 사항] 반영해서 v1.2 빌드해줘."

---

**마지막 작업 일자**: 2026-05-20
**마지막 커밋**: `claude/build-ir-deck-ri0ZR` 브랜치, `[v0.7] 또와 IR 덱 최종 정리 — Pre-Seed 5억 + Market Insight 신설`
