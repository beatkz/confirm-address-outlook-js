# ================================================
# Office Add-in Bootstrap Script (PowerShell)
# ================================================


Param(
    [switch]$debug,
    [switch]$halt
    )

# スクリプトのディレクトリに移動
Push-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

if ($debug) {
    & npm run build:dev
    & npm run start
}
elseif ($halt) {
    & npm run stop
} else {
    Write-Host "office-addin-bootstrap: Please specify either -debug, or -halt." -ForegroundColor Yellow
}