@echo off
pushd "%~dp0"
setlocal enabledelayedexpansion
:: If your current git branch is dev, execute the three steps: build, merge to main, return to dev.
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

if "%BETA_MODE%" == "TRUE" (
  echo Running in beta branch mode...
  call npm run build:dev
  if !errorlevel! neq 0 (
    echo Build failed. Please check the errors above.
    pause
    exit /b !errorlevel!
  )
  echo Starting beta branch deployment...
  
  REM Proceed with file copy for beta deployment
  echo Copying files for beta deployment...
  del /q /s /f "dist\manifest.*"
  rd /s /q "..\docs\beta"
  xcopy /s /y /i dist "..\docs\beta"
  xcopy /s /y /i "..\notes\manual" "..\docs\manual"
  copy /y "..\README.md" "..\docs\beta"
  copy /y manifest_beta.xml "..\docs\beta\manifest_beta.xml"

  echo After you commit, hit any key to continue...
  pause .

  REM Check for uncommitted changes
  git status --porcelain | findstr . >nul
  if !errorlevel! equ 0 (
    echo Uncommitted changes detected in dev branch.
    echo Please commit or stash all changes before proceeding.
    pause
    exit /b 1
  )

  REM Switch to main branch
  echo Switching to main branch...
  git checkout main
  if !errorlevel! neq 0 (
    echo Failed to switch to main branch.
    pause
    exit /b !errorlevel!
  )

  REM Merge dev branch into main
  echo Merging dev branch into main...
  git merge dev --no-ff
  if !errorlevel! neq 0 (
    echo Merge failed. Possible conflicts detected.
    echo Please resolve conflicts manually and try again.
    echo Switching back to dev branch...
    git checkout dev
    pause
    exit /b !errorlevel!
  )

  REM Switch back to dev branch
  echo Switching back to dev branch...
  git checkout dev
  if !errorlevel! neq 0 (
    echo Failed to switch back to dev branch.
    pause
    exit /b !errorlevel!
  )
) else (
  echo Running in release branch mode...
  call npm run build:dev
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