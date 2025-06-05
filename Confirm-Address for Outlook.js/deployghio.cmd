@echo off
setlocal enabledelayedexpansion
:: If your current git branch is dev, set BETA_MODE to TRUE.
REM get current git branch name
for /f "tokens=*" %%i in ('git branch --show-current') do set BRANCH=%%i

REM confirm that we are in a git repository
if "%BRANCH%"=="" (
    echo Could not determine the current git branch.
    echo Please ensure you are in a git repository.
    exit /b 1
)

REM display current branch name
echo Current branch name: %BRANCH%

IF "%BRANCH%" == "dev" (
  set BETA_MODE=TRUE
) ELSE (
  set BETA_MODE=FALSE
)

pushd "%~dp0"
if "%BETA_MODE%" == "TRUE" (
  echo "Running in beta branch mode..."
  call npm run build:dev
) else (
  echo "Running in release branch mode..."
  call npm run build
)
if %errorlevel% neq 0 (
  echo "Build failed. Please check the errors above."
  pause .
  exit /b %errorlevel%
)
if "%BETA_MODE%" == "TRUE" (
  echo "Starting beta branch deployment..."
  del /q /s /f "dist\manifest.*"
  rd /s /q "..\docs\beta"
  xcopy /s /y /i dist "..\docs\beta"
  xcopy /s /y /i "..\notes\manual" "..\docs\beta\manual"
  copy /y "..\README.md" "..\docs\beta"
  copy /y manifest_beta.xml "..\docs\beta\manifest_beta.xml"
) else (
  echo "Starting release branch deployment..."
  xcopy /s /y /i dist "..\docs"
  xcopy /s /y /i "..\notes\manual" "..\docs\manual"
  copy /y "..\README.md" "..\docs"
)

timeout 10
popd
endlocal