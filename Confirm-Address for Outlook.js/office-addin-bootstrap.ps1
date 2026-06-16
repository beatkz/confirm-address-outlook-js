# ================================================
# Office Add-in Bootstrap Script (PowerShell)
# ================================================


Param(
    [switch]$debug,
    [switch]$deploy,
    [switch]$halt
    )

# スクリプトのディレクトリに移動
Push-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

if ($debug) {
    & npm run build:dev
    & npm run start
}
elseif($deploy) {
    # 現在のGitブランチを取得
    $branch = git branch --show-current
    # ローカルモード（devブランチ）の判定
    $devMode = ($branch -eq "dev")

    if ($devMode) {
        Read-Host "You are currently on the 'dev' branch. After committing your changes, hit any key to continue"
        & git checkout main
    }
    & ./deployghio.ps1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build failed. Please check the errors above." -ForegroundColor Red
        exit $LASTEXITCODE
    }
    Write-Host "Build success. You can now deploy the add-in to the Microsoft Store or distribute it as needed." -ForegroundColor Green
}
elseif ($halt) {
    & npm run stop
} else {
    Write-Host "office-addin-bootstrap: Please specify either -debug, -deploy or -halt." -ForegroundColor Yellow
}