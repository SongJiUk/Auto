# 서식 채우기 프로그램

공사 현장의 공사일보·TBM일지·사진대지 등을 매일 자동으로 채워주는 Windows 프로그램.

---

## 프로그램 쓰는 방법 (회사·집 PC 공용) — 구글 드라이브

프로그램과 설정은 **구글 드라이브 폴더 하나**에 두고 두 PC가 같이 쓴다.  
GitHub에는 **코드만** 올린다.

| 무엇 | 어디에 | 비고 |
|------|--------|------|
| 실행 파일, 설정, 서식 | 구글 드라이브 `G:\내 드라이브\FormFiller\` | 자동 동기화 |
| 프로그램 코드 | GitHub (이 저장소) | 고칠 때 push / pull |

```
G:\내 드라이브\FormFiller\
  FormFiller.App.exe   ← 실행 (바탕화면 바로가기로)
  config.json          ← 설정 전부 (exe 옆에서 읽고 저장)
  서식\*.xlsx          ← 서식 파일 (설정에는 "서식\파일명" 으로 저장됨)
  output\백업\         ← 같은 날짜로 다시 만들 때 덮어쓰기 전 파일 (30일 보관)
  logs\log_PC이름.txt  ← 오류·생성 기록 (PC마다 따로)
```

### 새 PC 준비
1. 구글 드라이브 PC 프로그램 설치 → 로그인
   - 다른 계정이면: 원래 계정에서 `FormFiller` 폴더를 **편집자**로 공유 → 새 계정에서 **내 드라이브에 바로가기 추가**
2. `G:\내 드라이브\FormFiller\` 가 내려왔는지 확인 (exe가 커서 몇 분 걸릴 수 있음)
3. `FormFiller.App.exe` 를 바탕화면에 바로가기로 만들어 실행

### 지킬 것
- 두 PC에서 **동시에** 켜 두고 설정을 바꾸지 않는다 (나중 저장이 덮어씀).
- 다른 PC에서 설정을 바꿨으면 드라이브 동기화가 끝난 뒤 프로그램을 켠다.
- 새 서식은 설정 창에서 등록하면 `서식\` 폴더로 자동 복사된다.
- API 키를 넣으면 `config.json` 에 저장된다. 이 폴더는 남에게 공유하지 않는다.
- 구글 드라이브 설정은 **미러링** 이 좋다. 스트리밍이면 `FormFiller` 폴더와 사진 폴더를 **오프라인으로 사용 가능** 으로 해 둔다 (사진을 열 때마다 내려받아 느려짐).
- 이상하면 `logs\log_PC이름.txt` 를 확인한다.

### 프로그램을 고쳤을 때
코드 수정 → GitHub push → 새 exe를 `G:\내 드라이브\FormFiller\FormFiller.App.exe` 에 덮어쓰기.  
(덮어쓰기 전에 프로그램을 닫는다. 다른 PC에는 드라이브가 자동으로 전달한다.)

---

## 새 PC에서 개발 환경 세팅하는 방법 (코드를 고칠 때만)

### 1단계 — .NET 8 SDK 설치

아래 주소에서 **.NET 8 SDK** (Windows x64) 를 받아 설치한다.

```
https://dotnet.microsoft.com/download/dotnet/8.0
```

설치 후 터미널(cmd 또는 PowerShell)에서 확인:

```
dotnet --version
```

`8.x.x` 가 출력되면 OK.

---

### 2단계 — 저장소 클론

```
git clone https://github.com/<계정명>/<저장소명>.git
cd <저장소명>
```

---

### 3단계 — 서식 파일 복사

`.gitignore` 에 의해 `samples/` 폴더와 엑셀 파일은 저장소에 포함되지 않는다.  
기존 PC 또는 클라우드(OneDrive 등)에서 아래 파일들을 `samples/` 폴더로 복사한다.

```
samples/
  ├── 공사일보_출력일보_*.xlsx
  ├── TBM일지_*.xlsx
  └── 사진대지_*.xlsx   (있는 경우)
```

---

### 4단계 — 빌드 & 배포

```
dotnet publish src/FormFiller.App/FormFiller.App.csproj ^
  -c Release -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true -o publish
```

> PowerShell 에서는 `^` 대신 `` ` `` (백틱) 사용.

---

### 5단계 — 실행

```
publish\FormFiller.App.exe
```

또는 `publish` 폴더의 `FormFiller.App.exe` 를 더블클릭.

---

### 6단계 — 설정

앱을 처음 실행하면 `config.json` 이 없으므로 빈 상태다.  
상단 **설정** 버튼 → **+ 문서 추가** 로 서식 파일을 하나씩 등록한다.

> `config.json` 을 기존 PC에서 복사해 오면 설정을 다시 할 필요가 없다.  
> 위치: `publish\config.json`

---

## 개발 환경 (코드 수정 시)

| 항목 | 버전 |
|------|------|
| .NET SDK | 8.0 |
| 언어 | C# 12 |
| UI | WPF (.NET 8) |
| 엑셀 | DocumentFormat.OpenXml (Open XML SDK) |
| IDE | Visual Studio 2022 또는 VS Code + C# Dev Kit |

```
# 빌드만 확인
dotnet build

# 테스트
dotnet test

# 실행 (디버그)
dotnet run --project src/FormFiller.App
```

---

## 폴더 구조

```
서식채우기/
├── src/
│   ├── FormFiller.Core/   # 엑셀 읽기·쓰기 엔진 (화면 없음)
│   └── FormFiller.App/    # WPF 화면
├── docs/                  # 기획 문서
├── samples/               # 서식 파일 (git 제외 — 직접 복사)
├── output/                # 생성된 파일 (git 제외)
└── publish/               # 배포 실행파일 (git 제외)
```
