# Opptx 작업 일지

프로젝트 진행 상황, 릴리즈 이력, 남은 작업을 기록합니다. 새 작업을 마치면 맨 위 "작업 기록"에 날짜순으로 추가합니다.

- 저장소: https://github.com/moonlightpark/Open-Power-Point-Editor
- 현재 버전: **0.1.2** (`package.json`, `src-tauri/Cargo.toml`, `src-tauri/tauri.conf.json`, 정보 대화상자에서 함께 올립니다)
- 라이선스: MIT

## 현재 상태 요약

| 항목 | 상태 |
|---|---|
| 웹 편집기 (`npm run dev`, :5174) | 동작 |
| macOS 데스크톱 앱 (Apple Silicon / Intel) | 빌드, 실행 확인 |
| Windows / Linux 앱 | CI에서 빌드만 확인. 실제 설치·실행은 아직 시험하지 않음 |
| 엔진 테스트 | vitest 10개 통과 |
| 최신 릴리즈 | v0.1.2 (Latest, 공개) |

## 작업 기록

### 2026-10-16

- [x] `job_log.md` 추가 (프로젝트 진행 관리)

### 2026-10-16 (최초 개발과 릴리즈)

- [x] REA로 HOP 0.4.5 분석 → [docs/HOP-analysis.md](docs/HOP-analysis.md)
- [x] PPTX 엔진(OPC 패키지, 테마, 도형, 텍스트, 렌더러, undo)과 HOP 스타일 편집기 UI 구현 — `3d26bf1`
- [x] Tauri 2 데스크톱 앱 빌드 (Tauri 2.11 고정, 로컬 rustc 1.86)
- [x] 왼쪽 슬라이드 목록: 추가 · 삭제 · 복사/잘라내기/붙여넣기 · 복제, 오른쪽 클릭 메뉴, 단축키
- [x] MIT 라이선스, THIRD_PARTY_NOTICES.md(`npm run notices`), HOP·rhwp 참고 출처, 상표 고지 — `1fa0809`
- [x] GitHub Actions 릴리즈 워크플로(Windows, Linux, macOS ×2) — `53202e0`, `a60562f`
- [x] 새 앱 아이콘 — `7f4ecb5`
- [x] 릴리즈 초안이 중복 생성되던 문제 수정(create-release 잡) — `200e7c4`

## 릴리즈 이력

| 버전 | 날짜 | 내용 | 비고 |
|---|---|---|---|
| v0.1.2 | 2026-10-16 | 새 앱 아이콘, 전 플랫폼 빌드 | Latest. 자산 9개 |
| v0.1.1 | 2026-10-16 | Windows·Linux 빌드 추가, MIT 라이선스 | |
| v0.1.0 | 2026-10-16 | 최초 릴리즈 (macOS만) | 라이선스 고지 파일 없음 |

릴리즈 방법: 세 파일의 버전을 올리고 커밋한 뒤 `git tag vX.Y.Z && git push origin vX.Y.Z` → Actions가 초안을 만들고 빌드 결과를 올림 → 확인 후 공개.

## 할 일

- [ ] Windows / Linux 설치 파일을 실제 환경에서 실행해 보기
- [ ] 다음 태그에서 수정한 릴리즈 워크플로(초안 1개만 생성)가 제대로 도는지 확인
- [ ] v0.1.0 릴리즈 정리 여부 결정 (라이선스 고지가 빠진 빌드)
- [ ] 저장소 이름에서 "PowerPoint"를 뺄지 결정 (상표 관련)

## 개선 후보 (백로그)

- 차트, SmartArt 렌더링 (지금은 자리 표시 상자만 보이고, 저장할 때는 보존됨)
- 애니메이션, 전환 효과 재생
- 텍스트 줄바꿈을 PowerPoint와 더 가깝게 맞추기
- 코드 서명(macOS 공증, Windows 서명)

## 참고 메모

- Tauri 크레이트와 npm `@tauri-apps/api`, `@tauri-apps/cli`는 major.minor를 맞춰야 합니다(현재 2.11).
- 의존성이 바뀌면 `npm run notices`로 THIRD_PARTY_NOTICES.md를 다시 만듭니다.
- DMG 빌드가 실패하면 남아 있는 마운트 볼륨(`/Volumes/Opptx*`)을 먼저 확인합니다.
