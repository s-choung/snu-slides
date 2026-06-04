# snu-slides 스킬 — HANDOVER (세션 인계)

> 이 문서 + `SKILL.md` + `DESIGN_SYSTEM.md`만 읽으면 이어받을 수 있게 작성.
> 폴더명은 레거시로 `snu-slides`지만, **테마는 skku / yu / jnu 3교**다(서울대 SNU 아님).

## 0. 현재 구조 (`~/.claude/skills/snu-slides/`)
- `SKILL.md` — 진입점. 멀티스쿨(skku/yu/jnu) 테마 표, slidespec, 모드, **MANDATORY 검증 규칙**.
- `DESIGN_SYSTEM.md` — SKKU 레퍼런스 덱의 실측 토큰(SSOT). navy/gold 컴포넌트 치수·레시피.
- `themes/skku.json` (+`skku_components.xml`) · `themes/yu.json` · `themes/jnu.json` — 테마 토큰. yu/jnu엔 `official_ci` 블록(공식색·로고 URL·폰트).
- `assets/logos/{skku,yu,jnu}.png` — 표지용 투명 로고.
- `scripts/generate.py` — slidespec → HTML 덱. (레퍼런스, 잘 나옴)
- `scripts/export_pptx.py` — slidespec → .pptx. **skku는 원본 shape 클론, yu/jnu는 primitive fallback.**
- `scripts/clone.py` — `<theme>_components.xml`의 원본 `<p:sp>`를 `parse_xml`로 떠와 위치/텍스트만 치환. set_xfrm/set_text/set_number/set_markup/set_nowrap/set_solid_fill/dedupe_ids.
- `scripts/audit.py` — HTML+PPT per-slide PNG(≤1280px) 렌더 + manifest. **검증용.**
- `scripts/showcase.py` — 여러 학교 예시 덱을 한 페이지(히어로/갤러리)로. `showcase.py OUT DECKS skku yu jnu`.
- `scripts/build_deck.py`,`convert.py` — PPT→HTML 충실 복제(별개 도구, CONVERTER.md).

## 1. ★완료된 핵심 작업 (회귀 금지)
- **PPT 클론 전환(skku)**: primitive 드로잉 → 원본 OOXML 클론. **gold 알약칩이 PowerPoint에서 안 보이던 문제 해결**(원본 `roundRect adj=50000` shape 그대로). 카드 둥글기(adj=2939, r≈11)·3px 테두리·번호블록 전부 원본 일치.
- **번호블록 "0N" 렌더 버그**: 원본 numtext의 투명 텍스트 외곽선(`<a:ln alpha=0>`)이 soffice width 오측정 → 둘째 자리 클리핑. `set_number`가 외곽선 제거 + `set_nowrap`로 해결(앞0 흐림+solid 자리 유지).
- **HTML↔PPT 정합**: 배너를 원본 스택 레이아웃(라벨 위·칩 x=129 아래·제목 x=206)으로 통일. 카드 테두리 accent별 색. 번호블록 앞자리 흐림. 검증 clean.
- **멀티스쿨**: snu→skku 정정. skku는 **공식 CI(Navy #072B61 + Sungkyun Green #8DC63F + Orange #FF6C0F)**. 클론 컴포넌트는 원본 navy/gold 하드코딩이지만 `clone.recolor`로 테마색에 리매핑(형태 유지·색만 교체). 가독성: 칩 글자=navy(chip_text), 강조 텍스트/카드헤드=DeepGreen #124633(em_color). yu(blue #153974/#00AACA)·jnu(green #007A33/navy #003594)도 동일 구조(primitive).
- **표지(cover=제목 페이지)**: 로고(좌상단, theme `logo`) + 큰 제목(primary) + accent 구분선 + 부제 + 하단 primary 띠(presenter 좌/ date 우, 흰색). HTML+PPT 동일. spec 필드: `presenter`,`date`(legacy `meta` fallback).
- **쇼케이스 랜딩**: showcase.py로 3교 카드(로고·스와치·HTML/PPTX 링크) 생성. 검증 PASS.

## 2. 검증 규칙 (MANDATORY — SKILL.md에도 명시)
- **메인 에이전트는 슬라이드 PNG 직접 Read 금지**(컨텍스트 포화). 항상 서브에이전트 위임.
- 생성 후 `audit.py deck.html deck.pptx --out /tmp/_x` → manifest를 서브에이전트에 주고 (1)HTML 결함 (2)HTML↔PPT 동일성 텍스트 리포트 → clean까지 수정·재감사 → 그 다음 사용자에게.
- **soffice 전용 아티팩트**(실제 PowerPoint는 정상): 한글/라틴·구두점 경계 띄어쓰기("LLM 의", "요약 · 분류"), HTML 내비칩. 실제 PowerPoint 렌더는 별도 확인 권장(headless 불가).

## 3. 알려진 한계 / 다음 후보
- yu/jnu는 클론 소스(원본 덱)가 없어 primitive. 더 정교하게 하려면 학교별 reference pptx 확보 후 `<theme>_components.xml` 추출.
- yu/jnu 공식 폰트(윤고딕/Lucida, KoPubWorld/전남대체)는 미설치 → Arial+Malgun fallback. 필요시 폰트 설치 후 테마 fonts 갱신.
- 폴더명 `snu-slides`는 레거시. 원하면 `uni-slides` 등으로 리네임 가능(스킬 등록 영향 주의).
- skku는 공식 CI로 리컬러됨(레퍼런스 덱의 navy/gold #0F0F70/#C5A86F는 `clone.recolor` 매핑의 SOURCE 키일 뿐, 출력색은 공식 #072B61/#8DC63F). DESIGN_SYSTEM의 치수는 그대로 유효(형태), 색만 themes/skku.json이 지배.

## 4. 예시 산출물 (영구 보관: playground `121_ppt2html/`)
- 위치: `…/Research_2026/playground/121_ppt2html/` (스킬 폴더가 아님 — 사용자 지정).
- `index.html` — 3교 쇼케이스 랜딩. `{skku,yu,jnu}_deck.html`/`.pptx` — 표지+content+cards 3장 덱. `specs/{skku,yu,jnu}_deck.json` — 재현용 slidespec.
- 재현: `export_pptx.py 121_ppt2html/specs/skku_deck.json out.pptx`, `showcase.py OUT DECKS_DIR skku yu jnu`.
