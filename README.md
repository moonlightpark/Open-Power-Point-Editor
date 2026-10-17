# Opptx — PPTX 편집기

[HOP](https://github.com/golbin/hop)(HWP 편집기)의 구조를 REA로 분석해서 같은 방식으로 만든 PPTX 프레젠테이션 편집기입니다.
구조는 **문서 엔진, 렌더러, 명령 기반 UI, 데스크톱 셸**의 네 층으로 되어 있습니다. 분석 내용은 [docs/HOP-analysis.md](docs/HOP-analysis.md)에 있습니다.

## 설치

[Releases](https://github.com/moonlightpark/Open-Power-Point-Editor/releases)에서 운영체제에 맞는 파일을 받습니다.

| 운영체제 | 파일 | 참고 |
|---|---|---|
| Windows 10/11 (x64) | `Opptx_x.y.z_x64-setup.exe` 또는 `.msi` | 서명하지 않은 빌드입니다. SmartScreen 경고가 뜨면 **추가 정보 → 실행**을 누릅니다. WebView2가 없으면 설치 중에 자동으로 받습니다. |
| macOS (Apple Silicon) | `Opptx_x.y.z_aarch64.dmg` | 임시 서명(ad-hoc)이라 처음 실행할 때 앱을 **우클릭 → 열기**로 실행합니다. |
| macOS (Intel) | `Opptx_x.y.z_x64.dmg` | 위와 같습니다. |
| Linux (x64) | `.AppImage`, `.deb`, `.rpm` | WebKitGTK 4.1이 필요합니다(Ubuntu 22.04 이상, Fedora 36 이상 등). AppImage는 `chmod +x` 후 실행합니다. |

릴리즈 빌드는 `v*` 태그를 푸시하면 GitHub Actions([`.github/workflows/release.yml`](.github/workflows/release.yml))가 자동으로 만듭니다.

## 실행

```bash
npm install
npm run dev          # http://localhost:5174 (샘플: /?file=./sample.pptx)
npm test             # 엔진 테스트 (vitest + jsdom)
npm run build        # dist/ 웹 빌드

npx tauri dev        # 데스크톱 앱(Tauri 2) 개발 실행
npx tauri build      # .app / .dmg 번들
```

## 기능

- **열기/저장**: .pptx 열기(대화상자, 끌어다 놓기), 저장, 다른 이름으로 저장, 인쇄/PDF
  - 손대지 않은 XML 파트, 미디어, 차트, 애니메이션은 바이트 그대로 보존됩니다.
- **렌더링**: 마스터 → 레이아웃 → 슬라이드 상속(위치, 글꼴, 크기, 글머리표), 테마 색과 글꼴(`lumMod`/`lumOff`/`tint`/`shade`), 배경(단색, 그라데이션, 그림, `bgRef`), 프리셋 도형 20여 종과 사용자 지정 도형(`custGeom`), 선과 화살표 끝, 점선, 그림(자르기, 대칭), 그룹, 표(병합, 테두리, 표 스타일), 자동 맞춤, 세로 텍스트
- **편집**
  - 개체 선택(클릭, Shift, 드래그 범위), 이동, 크기 조절, 회전(Shift 스냅), 스마트 가이드 스냅, 방향키 미세 이동
  - 텍스트 직접 편집(더블클릭, Enter, F2): 글꼴, 크기, 진하게, 기울임, 밑줄, 취소선, 색, 정렬, 글머리표, 번호, 수준(Tab)
  - 바뀐 서식만 XML에 쓰기 때문에 placeholder 상속은 유지됩니다.
  - 표 셀 편집, 빈 placeholder 안내 문구
  - 삽입: 글상자, 도형, 그림(파일, 붙여넣기, 끌어다 놓기), 표
  - 채우기, 선 색, 개체 속성(cm 단위 위치·크기·회전), 순서(맨 앞/맨 뒤), 맞춤, 대칭
  - 슬라이드: 새로 만들기(레이아웃 선택), 복제, 삭제, 끌어서 순서 바꾸기, 레이아웃 변경, 배경 색
  - 왼쪽 슬라이드 목록: 위쪽 버튼(새 슬라이드 · 복제 · 삭제), 미리보기에 마우스를 올리면 나오는 복제/삭제 버튼, 슬라이드 사이의 `+`로 그 자리에 추가, 오른쪽 클릭 메뉴, 단축키(Ctrl+C/X/V/D, Delete, Enter, Ctrl+A). Shift/Ctrl 클릭으로 여러 장을 한꺼번에 처리하며, 복사한 슬라이드는 원본을 지운 뒤에도 붙여넣을 수 있음
  - 복사/붙여넣기(슬라이드 사이에서도 그림 관계 유지), 찾기/모두 바꾸기, 되돌리기/다시 실행(200단계)
- **페이지 형식**: 새 문서는 **A4 가로(29.7 × 21 cm)** 로 만들어져 인쇄할 때 용지에 꼭 맞습니다.
  - 파일 → 페이지 설정(또는 상태바의 형식 이름 클릭)에서 A4 가로 · A4 세로 · **웹 페이지(16:9)** · 화면(4:3)으로 바꿀 수 있습니다.
  - 바꿀 때 내용은 비율 유지하며 맞추기 / 페이지에 맞게 늘리기 / 그대로 두기 중에서 고릅니다. 되돌리기(Ctrl+Z)도 됩니다.
  - 툴바의 **가로 / 세로** 버튼(또는 슬라이드 메뉴)으로 모든 슬라이드의 방향을 한 번에 바꿉니다. PPTX는 문서 전체에 슬라이드 크기를 하나만 저장하므로 방향도 문서 단위로 정합니다.
  - 도구 → 옵션에서 새 문서의 기본 형식을 정합니다(이 브라우저/앱에 저장).
- **보기**: 확대/축소(Ctrl+휠), 눈금자(cm), 격자, 슬라이드 노트, 슬라이드 쇼(F5)

## 구조

```
src/engine/          문서 엔진 (UI 의존성 없음)
  package.ts         OPC 컨테이너: zip, 관계, 콘텐츠 형식, undo용 스냅샷
  document.ts        PptxDocument: 슬라이드/도형/서식 명령 + 트랜잭션 undo
  render.ts          슬라이드 → DOM/SVG 렌더러 (1px = 1pt = 12700 EMU)
  text.ts            텍스트 스타일 상속, 렌더링, 편집한 DOM → <a:p> 변환
  theme.ts           테마 색과 글꼴, 색 변환
  geometry.ts        프리셋/사용자 지정 도형 경로
src/ui/              편집기 UI (HOP와 같은 메뉴/툴바/서식 바/눈금자/상태바)
src-tauri/           데스크톱 셸 (네이티브 파일 대화상자, 임시 파일 후 rename 방식 저장)
tests/               엔진 테스트 + 샘플 덱
```

## 제약

- 차트와 SmartArt는 자리 표시 상자로만 보입니다(저장할 때는 보존됨).
- 애니메이션과 전환 효과는 재생하지 않습니다(보존됨).
- 텍스트 줄바꿈은 브라우저 레이아웃을 쓰기 때문에 PowerPoint와 약간 다를 수 있습니다.

## 참고한 프로젝트

이 프로젝트는 **[HOP](https://github.com/golbin/hop)** (golbin/hop, MIT License)을 참고했습니다.

- HOP은 [rhwp](https://github.com/edwardkim/rhwp) (MIT License) 엔진을 쓰는 HWP 문서 편집기입니다. Tauri 데스크톱 셸, Vite 웹 UI, Rust/WASM 문서 엔진으로 되어 있습니다.
- HOP 0.4.5 macOS 앱을 [REA](https://github.com/morluto/rea)로 분석해서 다음을 참고했습니다. 분석 내용은 [docs/HOP-analysis.md](docs/HOP-analysis.md)에 있습니다.
  - 문서 엔진, 렌더러, 명령 기반 UI, 데스크톱 셸로 나누는 구조
  - 메뉴바, 아이콘 툴바, 서식 바, 눈금자, 상태바로 이루어진 화면 배치와 그 요소 이름(`#menu-bar`, `#icon-toolbar`, `#style-bar`, `#status-bar` 등)
  - `category:action` 형식의 명령 체계와, 임시 파일에 쓴 뒤 바꿔치는 안전한 저장 방식
- HOP과 rhwp의 소스 코드는 포함하거나 복사하지 않았습니다. Opptx의 PPTX 엔진, 렌더러, UI는 모두 새로 작성했습니다.

HOP과 rhwp 개발자분들께 감사드립니다.

## 라이선스

[MIT License](LICENSE) © 2026 moonlightpark

배포 앱에 포함된 서드파티 소프트웨어(npm 패키지 14개, Rust 크레이트 248개)의 라이선스는 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)에 있습니다. 의존성이 바뀌면 `npm run notices`로 다시 만듭니다.
선택할 수 있는 라이선스(예: JSZip의 `MIT OR GPL-3.0-or-later`)는 MIT를 골라 사용합니다.

## 상표

Opptx는 Microsoft Corporation과 관계가 없고, Microsoft가 보증하지 않습니다.
Microsoft와 PowerPoint는 Microsoft Corporation의 상표입니다.
PPTX 형식은 공개 표준인 ECMA-376 / ISO/IEC 29500(Office Open XML)에 따라 구현했습니다.
