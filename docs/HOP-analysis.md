# HOP 분석 노트 (REA)

대상: `/Applications/HOP.app` — HOP 0.4.5 (`net.golbin.hop`), arm64 Mach-O, 50 MB
소스: https://github.com/golbin/hop (MIT License). 엔진은 https://github.com/edwardkim/rhwp (MIT License)
분석 도구: REA (`inspect_rust_binary`, `open_binary`) 와 정적 문자열 및 리소스 추출

## 관찰한 내용 (바이너리에서 직접 확인)

| 항목 | 근거 |
|---|---|
| 데스크톱 셸은 **Tauri 2.10.3**, 웹뷰는 wry 0.54 / WKWebView | 바이너리에 남은 `tauri-2.10.3/src/...`, `wry-0.54.4/src/wkwebview` 경로 |
| 플러그인: dialog 2.7, fs 2.5, store 2.4, log 2.8 | `tauri-plugin-*` 크레이트 경로 |
| 문서 엔진은 **rhwp**(Rust). 파서, 레이아웃, 렌더러, 명령을 포함 | `third_party/rhwp/src/{parser,renderer,document_core/commands,...}` 경로 |
| 엔진을 **WASM으로 빌드해 웹 UI에서 실행** | 임베디드 에셋 `assets/rhwp_bg-*.wasm` (11.8 MB) |
| 프런트엔드는 Vite 번들(`index-*.js` 2 MB, `index-*.css`) | Tauri 임베디드 에셋(brotli)에서 추출한 `index.html` |
| Rust 쪽 명령(`invoke`)은 파일 수명주기만 담당 | `prepare_document_open`, `prepare_staged_hwp_save`, `commit_staged_hwp_save`, `export_pdf_from_hwp_path`, `list_recent_documents`, `mark_document_dirty`, `check_external_modification`, `create_editor_window` 등 |
| WASM 문서 API는 468개 메서드(`HwpDocument`) | `insertText`, `deleteText`, `applyCharFormat`, `applyParaFormat`, `createTable`, `renderPageToCanvas`, `getCursorRect`, `getSelectionRects`, `exportHwp`, `exportHwpx`, `beginBatch`/`endBatch` … |
| UI 구조 | `#menu-bar` (파일·편집·보기·입력·서식·쪽·표·도구, `data-cmd="category:action"`), `#icon-toolbar` (`.tb-group` / `.tb-btn`), `#style-bar` (글꼴, 크기, 가/가/가/가, 색, 정렬), `#editor-area` (`#h-ruler`, `#v-ruler`, `#scroll-container`), `#status-bar` (쪽, 구역, 모드, 확대) |

에셋 추출 방법: `__DATA_CONST`에 있는 Tauri의 `(경로 &str, 데이터 &[u8])` 테이블을 체인드 픽스업 포인터로 따라가서 brotli로 압축을 풀었습니다.
스크립트는 세션 스크래치패드의 `tools/extract.py`에 있습니다.

## 추론한 내용

- 렌더링과 편집은 전부 WASM 엔진 안에서 일어나고, JS UI는 엔진의 명령을 부르고 결과를 캔버스에 그리는 얇은 층입니다.
- 저장은 "staged" 방식입니다. 임시 파일에 쓴 다음 커밋해서 원본이 손상되지 않게 합니다.

## Opptx에 옮긴 설계

| HOP | Opptx |
|---|---|
| rhwp 엔진 (Rust→WASM): 모델 + 명령 + 렌더 | `src/engine/` (TypeScript): `document.ts` (모델/명령/undo), `render.ts` (렌더러), `text.ts` (텍스트 스타일 상속, 편집 반영) |
| HWP 바이너리 모델 | OOXML 파트를 원본 그대로 보존(XML DOM이 단일 원본)해서 손대지 않은 내용은 그대로 왕복 |
| `renderPageToCanvas` | `renderSlide` → HTML/SVG DOM. 텍스트 편집에는 contenteditable 사용 |
| `data-cmd` 명령 라우팅 | `src/ui/commands.ts` (`category:action` + 단축키) |
| 메뉴바 / 아이콘 툴바 / 서식 바 / 눈금자 / 상태바 | `index.html` + `src/styles.css` (같은 id와 클래스 체계) |
| Tauri 셸 (`prepare_*`, staged save) | `src-tauri/` — `open_document`, `save_document`(임시 파일에 쓰고 rename), `read_document`, `launch_paths` |

## 분석하지 않은 부분

- rhwp 내부 레이아웃 알고리즘은 디컴파일하지 않았습니다. HWP 전용이라 PPTX에는 적용되지 않습니다.
- 런타임 관찰(실행 추적)은 하지 않았고, 정적 증거만 사용했습니다.
