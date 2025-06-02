pushd "%~dp0"
call npm run build
if %errorlevel% neq 0 (
  echo "Build failed. Please check the errors above."
  pause .
  exit /b %errorlevel%
)
rd /s /q "..\docs"
mkdir "..\docs"
xcopy /s /y /i dist "..\docs"
copy /y "..\README.md" "..\docs"
pause .
popd