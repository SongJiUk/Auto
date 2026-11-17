@echo off
chcp 65001 > nul
echo 빌드 중...

dotnet publish src/FormFiller.App/FormFiller.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:DebugType=none -o publish

if %errorlevel% neq 0 (
    echo.
    echo 빌드 실패!
    pause
    exit /b 1
)

echo.
echo 완료! publish\업무자동화.exe 에 생성되었습니다.
pause
