# 슬라이드 5 — 또와가 일하는 모습 / 컨셉 이미지 프롬프트

> IR 덱 v0.8 슬라이드 5 (신규) 에 들어갈 *제품 컨셉 이미지* 생성용 프롬프트 모음.
> 도구: DALL·E 3 / Midjourney v6 / SDXL 등 어디서든 쓸 수 있도록 정리.

---

## 1. 컨셉 한 줄

**사장님은 눈앞의 손님에게만 집중 — 또와(반투명 고스트 3마리)는 *동시에* 단골·재고·이탈을 챙기는 *보이지 않는 막내 직원*.**

핵심 메타포: *"있는 줄도 모르게 다 해주는 AI 직원."*

---

## 2. 구도 (Composition)

- **앵글**: 정육점 매장 내부, *측면 / 약간 비스듬한 3/4 뷰* — 사장님·손님·고스트가 한 프레임에 다 보임
- **전경 (Foreground)**:
  - **사장님** — 앞치마 입고 도마 앞에서 *고기를 썰며 손님을 응대 중*. 표정은 친근·집중. *고스트는 안 보임*
  - **손님** — 진열대 너머에서 사장님과 대화
- **중·후경 (Mid/Background)**:
  - **반투명 고스트 3마리** — 매장 안을 *분주히* 움직임. 사장님 옆/뒤/카운터 쪽 *각자 다른 위치*
  - 각 고스트 위에 *말풍선* 1개씩
- **배경 소품**: 진열대 안 고기 부위들(목살·삼겹살·등심), POS 단말, 종이 영수증, 칠판 메뉴판

---

## 3. 고스트 디자인

- **외형**: 둥글둥글한 *클래식 만화 고스트* (꼬리 흘러내리는 형태). 너무 무서운 X / 귀엽고 친근 O
- **투명도**: 50~60% 반투명 — 뒷배경 살짝 비침. *사장님·손님은 인지 못하는* 느낌 강조
- **색상**: AWS 오렌지 (#D15127) 계열 — 메인 컬러. 또는 *연한 오렌지* (#FAD9CB) 본체 + *진한 오렌지* (#A8401E) 외곽선
- **표정**: 각자 다른 *일하는 표정* — 휴대폰 보기 / 진열대 살피기 / 노트북·태블릿 들고 있기
- **소품**: 고스트마다 *작은 도구* — 휴대폰 / 클립보드 / 태블릿

---

## 4. 말풍선 (Speech Bubbles)

> 한국어. 손글씨 느낌 X / *깔끔한 산세리프* (Noto Sans KR 호환) O.
> 말풍선 자체는 *흰 배경 + 오렌지 테두리*.

| 고스트 | 위치 | 말풍선 텍스트 | 의미 매핑 |
|---|---|---|---|
| **🅐** | 사장님 옆 (POS 근처) | *"정육왕 고객님! 보낸 문자 보고 찾아주셨군요! 감사해요!"* | 단골관리 · 재방문 유도 |
| **🅑** | 진열대 위 (목살 위) | *"엇! 사장님, 목살 재고가 평소보다 많아요! 지금 프로모션 돌릴게요!"* | 재고 분석 · 자동 마케팅 |
| **🅒** | 카운터/계산대 뒤 | *"사장님! 오늘 기준 이탈 고객 N명이 예상되는데요. 쿠폰 보내볼까요?"* | 이탈 예측 · 자동 대응 |

---

## 5. 색상·스타일

- **메인 컬러**: AWS 오렌지 (#D15127) — 고스트·말풍선 강조선
- **보조 컬러**: 다크 네이비 (#232F3E) — 사장님 앞치마, 텍스트
- **배경**: 따뜻한 베이지 / 옅은 우드 톤 — 정육점 *친근한 동네 매장* 느낌
- **고기 묘사**: 자연스러운 *핑크·레드* — 너무 그래픽하게 빨갛지 않게
- **스타일 키워드**: *flat illustration*, *modern editorial*, *clean line art with soft fills*, *slight grain*, *no photo-realism*

> 톤: *Stripe / Linear / Notion 일러스트* 같은 *모던 B2B SaaS 일러스트* — 너무 만화적이지 않고, 너무 사실적이지도 않은 중간 톤.

---

## 6. 프롬프트 — 영어 (DALL·E 3 / Midjourney)

### 6.1 메인 프롬프트

```
Modern flat editorial illustration of a Korean traditional butcher shop interior, 3/4 angle view.
A butcher shop owner (apron, friendly expression) stands at a wooden cutting board in the foreground, slicing meat while talking to a customer across the display counter. Pork cuts (samgyeopsal, mok-sal) visible in the glass display case, POS terminal on the side, chalkboard menu on the back wall.
Around the owner, THREE SEMI-TRANSPARENT CARTOON GHOSTS (50% opacity, classic round ghost shape with a flowing tail) are floating and working busily — the owner and customer cannot see them.
- Ghost A near the POS, holding a smartphone, with a speech bubble.
- Ghost B above the meat display, looking at the inventory, with a speech bubble.
- Ghost C near the cash register holding a tablet, with a speech bubble.
All ghosts are in AWS orange tone (#D15127 outline, #FAD9CB fill) with friendly expressions.
Speech bubbles are white with orange outlines, clean sans-serif Korean text inside.
Color palette: AWS orange (#D15127), dark navy (#232F3E), warm beige background, natural pink-red meat tones.
Style: modern B2B SaaS editorial illustration (like Stripe, Linear, Notion), flat with soft fills, clean line art, slight grain texture, NO photo-realism, NO chibi/anime, NO 3D render.
16:9 aspect ratio, white background outside the scene.
```

### 6.2 영어 압축 버전 (Midjourney용)

```
Korean butcher shop interior, 3/4 view --- shop owner in apron slicing meat at wooden board, talking to customer at display counter --- THREE semi-transparent orange ghosts (50% opacity, classic round cartoon ghost shape) floating around the shop, working — one near POS with smartphone, one above meat display, one at cash register with tablet, each with a speech bubble in Korean --- AWS orange (#D15127) and dark navy (#232F3E) palette, warm beige background --- modern flat editorial illustration, Stripe/Linear/Notion style, clean line art with soft fills, no photo-realism, no anime --- 16:9
```

---

## 7. 프롬프트 — 한국어 (국내 도구 / 디자이너 브리프용)

```
[장면] 한국 동네 정육점 내부, 비스듬한 3/4 앵글.
- 전경: 정육점 사장님(앞치마, 친근한 표정)이 진열대 앞 도마에서 고기를 썰며 손님을 응대 중. 진열대 안에는 목살·삼겹살·등심 부위가 보임. 옆에 POS 단말, 뒤 벽에 칠판 메뉴판.
- 사장님·손님 주변에 반투명 만화 고스트 3마리(둥근 클래식 고스트 형태, 50% 투명도, AWS 오렌지 #D15127 외곽선 + 연한 오렌지 #FAD9CB 본체)가 분주히 움직임. *사장님·손님은 고스트를 인지하지 못함.*
  - 고스트 A: POS 옆, 휴대폰을 들고 있음. 말풍선: "정육왕 고객님! 보낸 문자 보고 찾아주셨군요! 감사해요!"
  - 고스트 B: 진열대 위(목살 부근), 재고를 살핌. 말풍선: "엇! 사장님, 목살 재고가 평소보다 많아요! 지금 프로모션 돌릴게요!"
  - 고스트 C: 계산대 뒤, 태블릿을 들고 있음. 말풍선: "사장님! 오늘 기준 이탈 고객 N명이 예상되는데요. 쿠폰 보내볼까요?"
- 말풍선은 흰 배경 + 오렌지 테두리, 깔끔한 산세리프 (Noto Sans KR 호환).

[색상]
- 메인: AWS 오렌지 #D15127
- 보조: 다크 네이비 #232F3E (사장님 앞치마, 텍스트)
- 배경: 따뜻한 베이지·우드 톤
- 고기: 자연스러운 핑크·레드

[스타일]
- 모던 B2B SaaS 에디토리얼 일러스트 (Stripe / Linear / Notion 톤)
- 플랫 + 부드러운 채색 + 깔끔한 선화 + 약간의 노이즈 텍스처
- 절대 NO: 포토리얼리즘, 치비·애니메, 3D 렌더
- 16:9 비율, 장면 바깥은 흰 배경
```

---

## 8. 변형(Variations) — 같은 컨셉 다른 컷

필요시 추가로 시도해볼 만한 변형:

1. **클로즈업 컷** — 사장님 어깨 너머로 고스트 1마리만 보이는 1/3 컷 (전체 컷의 디테일 강조용)
2. **단순화 컷** — 사람 없이 *정육점 풍경 + 고스트 3마리만* (아이콘적으로 가장 깔끔)
3. **타임라인 컷** — 매장 안 시간 흐름을 좌→우로, 고스트가 *아침→점심→저녁* 각자 다른 일을 하는 3컷 만화
4. **결과 컷** — 사장님 표정이 *"단골이 늘었네?"* 하고 미소짓는 닫기 컷 (시리즈 마지막용)

→ 1차 발표용으론 *메인 프롬프트 1컷*이면 충분. 2~3은 디자이너 작업 시 확장 옵션.

---

## 9. 체크리스트 — 시안 검수

이미지 받았을 때 확인할 항목:

- [ ] 고스트가 *반투명*인가 (불투명하면 사장님이 인지하는 듯한 인상)
- [ ] 사장님이 고스트를 *인지하지 않는* 표정인가 (눈맞춤 X)
- [ ] 말풍선 텍스트가 *정확*하고 *읽기 좋게* 들어갔는가 (한국어 폰트 깨짐 잦음)
- [ ] AWS 오렌지 컬러가 *너무 빨갛지/노랗지* 않은가 (#D15127 톤 일관)
- [ ] *너무 만화·치비*하지 않은가 (B2B SaaS 톤 유지)
- [ ] *너무 사실적*이지 않은가 (사진 같으면 부담스러움)
- [ ] 16:9 비율 + 슬라이드에 텍스트 올릴 *여백*이 있는가

---

## 10. 사용 흐름

1. *DALL·E 3* 또는 *Midjourney v6*에 §6.1 메인 프롬프트 입력
2. 4~6장 생성, 위 §9 체크리스트로 1차 컷 선별
3. 말풍선 텍스트가 *깨졌으면* (한국어 흔히 발생): 이미지에서 말풍선 영역만 *Figma/PPT에서 덮어 그리기* (말풍선 위치만 잡고 텍스트는 직접 입력)
4. PPT 슬라이드 5에 삽입 — 위치: 슬라이드 중앙 90% 차지, 위·아래 12% 마진 (제목·캡션 영역)

---

**최종 산출물**: `projects/butcher-crm/IR/assets/slide5_concept_image.png` *(미생성 — 외부 도구로 생성 후 저장)*
