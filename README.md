# 평가원 기출 검색 및 AI 윤문 
KICE2011~

## 배포 대상 (2026-11-12부터 둘)

| 대상 | 구글 시트 | 코드 | clasp |
| --- | --- | --- | --- |
| 기존 | 「기출과 윤문(2011~)」 | `kice-1.0.0` — 브랜치 `kice-1.x`, 태그 `kice-1.0.0` | 미설정(편집기에서 직접 반영) |
| 신규 | 「KICE 윤문」(Drive `KICE윤문/`) | `kice-2.0.0` — `main`에서 개발 | 로컬 `.clasp.json`(신규 파일의 바운드 프로젝트, gitignore — 다른 GAS 레포와 같은 관례) |

- **로컬 `.clasp.json`은 신규 파일만 가리킨다.** `main` = `kice-2.0.0`(2026-11-13 첫 push).
- push 전에 라이브를 임시 폴더로 pull해 레포와 대조한다.
- 기존 파일의 코드를 고칠 일이 생기면 브랜치 `kice-1.x`에서 고치고 편집기에 직접 반영한다.
- 재구축 계획: 문서 레포 `kice-proofread-docs`의 `plans/2026-11-12-kice-rebuild-plan-v5.md`.

## kice-2.0.0 (2026-11-13)

이 프로젝트는 LLM을 부르지 않는다. 윤문은 맥미니의 `kice_engine`이 하고, 여기는 시트 창구다.

| 파일 | 내용 |
| --- | --- |
| `RemoteApi.gs` | 웹앱 `doGet`·`doPost` → 한 디스패처. 토큰 인증 |
| `kiceAuto.gs` | 배치 상태와 명령: `start`(러너) · `claim`·`heartbeat`·`submit`·`finish`(엔진) · `status`·`stop`·`tick` |
| `kiceData.gs` | `corpus`·`prompts`·`image`·`dsread`·`corpus_put` |
| `batchKiceMulti.gs` | 메뉴 `▶ Que 윤문 요청`·`⏹ 중지` (요청만 적고 엔진이 처리) |
| `ArchieveViewer.gs`·`ArchiveViewer.html` | 아카이브 뷰어. 속성 `QV_LAYOUT_SWITCH_ROW_2`부터 F열 라벨 "합의" |
| `tests/` | 가짜 GAS 런타임과 계약 시나리오. `node tests/test.js .` |

- 러너가 읽는 `start`·`status`·`tick`·`stop` 응답 모양은 1.0.0과 같다.
- 스크립트 속성: `REMOTE_TOKEN` · `LATEX_SS_ID` · `LATEX_KICE_SS_ID` · `LEGACY_SS_ID` · (전환 때) `QV_LAYOUT_SWITCH_ROW_2` · `GITHUB_TOKEN`·`GITHUB_REPO`. LLM 키는 넣지 않는다.
- 편집기에서 `rapi_setupCheck`를 한 번 실행하면 권한 승인과 설정 점검이 된다.
- 아직 없는 것: 뷰어의 의견별 채택/기각 버튼(Feedback 기록).

