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
elseif ($deploy) {
    # 現在のGitブランチを取得
    $branch = git branch --show-current

    # ローカルモード（devブランチ）の判定
    $localMode = ($branch -eq "dev")

    if ($localMode) {
        Read-Host -Prompt "You are currently on the 'dev' branch. After committing your changes, hit any key to continue."
    }

    & git checkout main
    & .\deployghio.ps1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Deployment failed. Please check the errors above."
        exit $LASTEXITCODE
    }
    & git checkout dev
}
elseif ($halt) {
    & npm run stop
} else {
    Write-Host "office-addin-bootstrap: Please specify either -debug, -deploy, or -halt." -ForegroundColor Yellow
}