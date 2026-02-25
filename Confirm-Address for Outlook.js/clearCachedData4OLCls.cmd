pushd %LocalAppData%\Microsoft\Outlook

rd /s /q HubAppFileCache
rd /s /q RoamCache
del /q *.*

popd
pause.