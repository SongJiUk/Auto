---
name: deploy
description: 서식채우기(FormFiller) 프로그램을 빌드·테스트하고 GitHub에 push한 뒤, 사용자가 실제로 쓰는 구글 드라이브 폴더(G:\내 드라이브\업무자동화\)의 exe를 새 것으로 바꾼다. "배포", "드라이브에 올려줘", "새 exe로 바꿔줘", "/deploy" 요청 때 사용.
---

# 배포 (빌드 → 테스트 → push → 드라이브 exe 교체)

사용자가 매일 쓰는 프로그램을 바꾸는 작업이다.
**2026-11-22 사용자 결정: 테스트·빌드가 통과하면 push 와 드라이브 교체는 묻지 않고 진행한다.** 아래 "(확인 필요)" 단계는 이상이 있을 때만 멈춘다 (테스트·빌드 실패, 프로그램 실행 중, `*.xlsx`·개인정보가 커밋에 섞임, exe 이름 변경).
exe 는 드라이브에만 복사한다. 바탕화면에는 복사하지 않는다.
명령은 프로젝트 루트(`서식채우기\`)에서 PowerShell로 실행한다.

## 고정 정보
| 항목 | 값 |
|---|---|
| 실행 폴더 | `G:\내 드라이브\업무자동화\` (exe + `config.json` + `서식\` + `output\` + `logs\`) |
| exe 이름 | `src\FormFiller.App\FormFiller.App.csproj` 의 `<AssemblyName>` + `.exe` — **매번 csproj에서 읽는다** (2026-11-18 기준 `업무자동화`) |
| 절대 건드리지 않는 것 | 드라이브의 `config.json`, `config*백업*.json`, `서식\`, `output\`, `logs\` |

## 1. 사전 점검
```powershell
git status --short
git log origin/main..HEAD --oneline   # 아직 push 안 한 커밋
```
- 커밋 안 된 코드 변경이 있으면 목록을 보여 주고 "커밋할지 / 그대로 둘지" 묻는다. 직접 커밋하지 않는다.
- `publish/config.json` 은 git에 추적되고 있다(CLAUDE.md "config.json은 GitHub에 올리지 않는다"와 어긋남). 이 파일의 변경은 커밋에 넣지 말고 사용자에게 알린다.

## 2. 테스트 + 빌드
```powershell
dotnet test FormFiller.sln -c Release
dotnet publish src/FormFiller.App/FormFiller.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:DebugType=none -o publish
```
- `빌드.bat` 과 같은 명령이다(`.bat` 은 `pause` 때문에 직접 실행하지 않는다).
- 테스트 실패 또는 빌드 실패 → **여기서 멈추고** 오류 출력을 그대로 보고한다.
- 성공하면 `publish\<AssemblyName>.exe` 의 크기·수정 시각을 보여 준다.
- `IncludeNativeLibrariesForSelfExtract=true` 라서 exe 하나만 복사하면 된다. `publish\` 의 `*_cor3.dll` 은 예전 빌드 잔여물이다.

## 3. GitHub push (확인 필요)
- push할 커밋 목록을 보여 주고 사용자 확인 후 `git push origin main`.
- 저장소: `SongJiUk/Auto`. `*.xlsx`, `*.hwpx`, `samples/`, `output/` 이 올라가지 않는지 `git diff --stat origin/main..HEAD` 로 확인한다.

## 4. 드라이브 exe 교체 (확인 필요)
1. 실행 중인지 확인한다. 이름이 바뀐 적이 있으니 두 이름 모두 확인한다.
   ```powershell
   Get-Process -Name '업무자동화','FormFiller.App' -ErrorAction SilentlyContinue | Select-Object Name, Id, Path
   ```
   - 드라이브 경로의 프로세스가 있으면 **종료하지 말고** 사용자에게 프로그램을 닫아 달라고 한다. 다른 PC(회사/집)에서 실행 중일 수도 있다고 알린다.
2. 기존 exe를 백업한다: `G:\내 드라이브\업무자동화\<이름>_이전.exe` (덮어써도 되는 한 개만 유지).
3. `publish\<AssemblyName>.exe` → `G:\내 드라이브\업무자동화\` 복사.
4. 복사 후 두 파일의 SHA256이 같은지 확인한다 (`Get-FileHash`).

### exe 이름이 드라이브와 다를 때 (예: 드라이브에는 `FormFiller.App.exe`, 새 것은 `업무자동화.exe`)
- 바탕화면 바로가기는 예전 이름을 가리킨다. 사용자에게 아래 중 하나를 고르게 한다.
  - 새 이름으로 복사하고 바로가기 대상을 바꾼다 (예전 exe는 사용자가 확인 후 삭제)
  - 예전 이름으로 복사한다 (바로가기 그대로)
- `config.json` 은 exe 옆에서 읽으므로 이름이 바뀌어도 설정은 그대로 이어진다.

## 5. 보고
- 테스트 결과(통과/실패 개수), 빌드 결과, push한 커밋, 교체한 exe 경로·해시 일치 여부를 적는다.
- 사용자 체크리스트: ① 바로가기로 실행 ② 오늘 날짜로 문서 1개 생성 ③ 다른 PC에서 드라이브 동기화 후 실행.
- 실행하지 않은 단계는 "안 함"이라고 적는다.
