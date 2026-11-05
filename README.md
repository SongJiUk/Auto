# 서식 채우기 프로그램

공사 현장의 공사일보·TBM일지·사진대지 등을 매일 자동으로 채워주는 Windows 프로그램.

---

## 새 PC에서 처음 세팅하는 방법

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
