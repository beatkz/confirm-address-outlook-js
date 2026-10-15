#!/usr/bin/env pwsh
# ================================================
# Office Add-in Bootstrap Script (PowerShell)
# ================================================


Param(
    [switch]$debug,
    [switch]$daemon,
    [switch]$deploy,
    [switch]$halt,
    [switch]$help
    )

#簡易ヘルプの表示
if($help.IsPresent -or !$PSBoundParameters.Count) {
    Write-Host "office-addin-bootstrap:" -ForegroundColor Yellow
    Write-Host "usage: ./office-addin-bootstrap.ps1 -debug|-daemon|-deploy|-halt" -ForegroundColor Yellow
    Write-Host "-debug -> Build and start the add-in in development mode" -ForegroundColor Green
    Write-Host "-daemon -> Run the development server in the background(Ctrl+C to quit)" -ForegroundColor Green
    Write-Host "-deploy -> Deploy the add-in" -ForegroundColor Green
    Write-Host "-halt -> Stop the development server" -ForegroundColor Green
    exit 0
}

# スクリプトのディレクトリに移動
Push-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

if ($debug) {
    & npm run build:dev
    & npm run start
}
elseif($daemon) {
    & npm run dev-server
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
}