---
name: doc-register
description: 서식채우기 프로그램에 새 서식 문서(엑셀 xlsx 또는 한글 hwpx)를 등록하거나 기존 등록을 고친다 — config.json의 documents 항목, 필드 종류·칸 지정, 저장 이름 패턴, hwpx 자리 글자({{날짜}}) 점검. "문서 등록", "업무일지 등록", "작업계획서 등록", "새 서식 추가" 요청 때 사용.
---

# 문서 등록 (xlsx · hwpx)

## 원칙
- 칸 위치·시트 이름·저장 이름은 **config.json 에만** 넣는다. 코드에 고정하지 않는다.
- 실사용 설정은 `G:\내 드라이브\업무자동화\config.json` 이다. **직접 고치기 전에 반드시 사용자 확인**을 받고, 고치기 전 `config_백업_YYYYMMDD.json` 으로 복사한다. 프로그램(`업무자동화.exe` / `FormFiller.App.exe`)이 실행 중이면 닫은 뒤에 고친다 (닫을 때 프로그램이 덮어쓸 수 있음).
- 가능하면 사용자가 프로그램의 [새 문서 추가] 화면(`DocumentSetupWindow`, Excel에서 클릭 지정)으로 등록하게 안내한다. 손으로 JSON을 쓰는 건 화면으로 안 되는 경우만.
- 모르는 규칙(주말 묶음, 누계 이월 방식 등)은 추측하지 말고 묻는다. 이미 정한 것은 `docs/결정사항.md` 를 먼저 본다.

## 1. 서식 파일 받기
- 서식은 드라이브 `서식\` 폴더에 둔다. `templateFile` 은 `서식\파일명` **상대 경로**.
- `.hwp` 는 지원하지 않는다 → 사용자에게 한글에서 **다른 이름으로 저장 → hwpx** 를 요청한다.
- 서식 파일은 git에 올라가지 않는다(`*.xlsx`, `*.hwpx` 는 .gitignore에 있음).

## 2. 문서 항목 (`DocumentConfig`)
| 키 | 내용 |
|---|---|
| `id`, `name` | 고유 id, 화면에 보일 이름 |
| `templateFile` | `서식\…` |
| `outputPattern` | 저장 이름. 날짜 토큰 `{yy}`, `{MM}`, `{dd}` (예: `{yy}.{MM}.{dd} 공사일보(참마루건설).xlsx`) — 사용자가 쓰던 이름 그대로 |
| `sheetName` / `sheetNames` | 한 시트 / 같은 모양 여러 시트. 월이 들어간 시트 이름(`TBM일지 서식11월`)은 매달 바뀌는지 사용자에게 확인 |
| `frequency`, `frequencyDays` | 생성 주기 (`daily` 기본) |
| `usePreviousAsBase` | 앞 날짜 파일을 바탕으로 생성 (엑셀만, 기본 켜짐) |
| `includesSujangjae` | 수장재 인원 포함 여부 |
| `fields[]` | 아래 필드 |

필드 종류(`FieldConfig.type`): `date_korean`, `date_simple`, `text`, `number`, `multiline`, `input`, `photo`, `shape_position`, `equipment_table`, `personnel_table`, `active_equipment_names`, `tbm_personnel`, `attendance_rate`.
새 종류가 필요하면 엔진 코드 수정이 필요하다 → 계획을 먼저 보여 주고 승인 후 TDD로 진행한다.

## 3. 엑셀(xlsx) 등록
1. 칸 확인은 읽기만 한다 (openpyxl `read_only=True` 또는 `xlsx-check` 스크립트). 셀 값에 개인정보가 있으면 주소만 말한다.
2. 날짜 칸이 문자열인지 날짜 값인지 확인하고 맞는 필드 종류를 고른다.
3. 수식 칸은 필드로 지정하지 않는다 (엔진이 덮어쓰지 않음).
4. 등록 후 오늘 날짜로 한 번 생성해 `/xlsx-check` 로 검증한다.

## 4. 한글(hwpx) 등록
hwpx는 zip 안의 XML이다. 본문은 `Contents/section*.xml`.
1. 사용자가 한글에서 매일 바뀌는 자리에 `{{날짜}}`, `{{날씨}}` 처럼 자리 글자를 적는다 (`{{` `}}` 사이 40자 이내, `<>{}` 불가 — `HwpxWriter.PlaceholderRegex`).
2. **자리 글자가 쪼개졌는지 점검한다.** 한글은 글자 모양이 바뀌거나 고쳐 쓴 곳에서 `{{날`·`짜}}` 처럼 여러 `<hp:t>` 로 나눠 저장하기도 한다. 그러면 엔진이 못 찾는다.
   ```powershell
   python -X utf8 -I -c "import zipfile,re,sys; z=zipfile.ZipFile(sys.argv[1]); [print(n, sorted(set(re.findall(r'\{\{[^{}<>]{1,40}\}\}', z.read(n).decode('utf-8'))))) for n in z.namelist() if n.startswith('Contents/section')]" "<서식.hwpx>"
   ```
   - 사용자가 적었다는 자리 글자가 목록에 없으면 쪼개진 것이다 → 그 자리 글자를 지우고 **한 번에 다시 입력**해 달라고 요청한다.
3. 필드의 `placeholder` 를 비우면 `{{필드 label}}` 로 찾는다. 다르면 `placeholder` 에 정확히 적는다.
4. 생성 후 결과 hwpx에 `{{` 가 남아 있지 않은지 같은 명령으로 확인하고, 사용자에게 한글로 열어 확인을 요청한다(렌더링은 직접 못 봄).

## 5. 아직 정할 것 (CLAUDE.md 남은 할 일)
- 업무일지: 주말 묶음 파일 이름(`26.11.13~15`)은 "하루 한 파일" 규칙과 안 맞음 → 등록 전에 사용자와 정한다.
- 작업계획서: 자리 글자 ↔ 필드 연결표를 사용자와 먼저 만든다.

## 6. 보고
- 추가·수정한 문서 id와 필드 목록, 백업 파일 경로, 시험 생성 결과(`xlsx-check` 종료 코드 / hwpx 남은 자리 글자 수), 사용자 확인이 필요한 항목.
