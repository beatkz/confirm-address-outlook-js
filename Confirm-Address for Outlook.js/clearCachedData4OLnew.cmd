pushd %LocalAppData%\Microsoft\Outlook

rd /s /q HubAppFileCache
rd /s /q RoamCache

popd
pause.