@echo off
pushd "%~dp0"
setlocal enabledelayedexpansion
:: Due to webpack mode limitations, if your current git branch is dev, only build.
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
  set LOCAL_MODE=TRUE
) ELSE (
  set LOCAL_MODE=FALSE
)

if "%LOCAL_MODE%" == "TRUE" (
  echo Running in local deployment mode...
  call npm run build:dev
  if !errorlevel! neq 0 (
    echo Build failed. Please check the errors above.
    pause
    exit /b !errorlevel!
  ) else (
    echo Build success.
    echo - Run "npm run start" to start debugging.
    echo - Run "npm run stop" to stop debugging.
    echo - When debugging is complete, switch to the main branch and run this script again to deploy the release version.
  )
) else (
  echo Running in release branch mode...
  call npm run build
  if !errorlevel! neq 0 (
    echo Build failed. Please check the errors above.
    pause
    exit /b !errorlevel!
  )
  echo Starting release branch deployment...
  xcopy /s /y /i dist "..\docs"
  xcopy /s /y /i "..\notes\manual" "..\docs\manual"
  copy /y "..\README.md" "..\docs"
)

timeout 10
popd
endlocal