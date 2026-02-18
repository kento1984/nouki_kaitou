@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM nouki_kaitouパッケージを認識させるため親ディレクトリをPYTHONPATHに追加
set "PYTHONPATH=%~dp0..;%PYTHONPATH%"

echo SAPの受注一覧ファイルを選択してください...
for /f "delims=" %%i in ('powershell -ExecutionPolicy Bypass -File "_select_file.ps1"') do set "SELECTED_FILE=%%i"

if "%SELECTED_FILE%"=="" (
    echo キャンセルされました。
    pause
    exit /b 1
)

echo 選択ファイル: %SELECTED_FILE%
echo.
echo 納期回答書作成を起動中...
python -m nouki_kaitou.main --source "%SELECTED_FILE%"

if errorlevel 1 (
    echo.
    echo エラーが発生しました。
    pause
)
