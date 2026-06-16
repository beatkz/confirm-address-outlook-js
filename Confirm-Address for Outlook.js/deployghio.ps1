# ================================================
# Deployment Script (PowerShell)
# ================================================

# スクリプトのディレクトリに移動
Push-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

try {
    # ====================== 関数定義 ======================

    # フォルダの中身を構造を保持したままコピーする関数
    function Copy-Contents {
        param(
            [Parameter(Mandatory=$true)]
            [string]$SourcePath,
            
            [Parameter(Mandatory=$true)]
            [string]$DestinationPath,
            
            [string]$Description = ""
        )
        
        if (-not (Test-Path $SourcePath)) {
            Write-Warning "'$SourcePath' が見つかりません。スキップします。"
            return $false
        }
        
        # 宛先フォルダが存在しない場合は作成
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # 中身だけをコピー（構造保持）
        Copy-Item -Path "$SourcePath\*" -Destination $DestinationPath -Recurse -Force
        
        if ($Description) {
            Write-Host "✓ $Description" -ForegroundColor Gray
        } else {
            Write-Host "✓ Copied contents of '$SourcePath' to '$DestinationPath'" -ForegroundColor Gray
        }
        
        return $true
    }
    
    # 単一ファイルをコピーする関数
    function Copy-File {
        param(
            [Parameter(Mandatory=$true)]
            [string]$SourceFile,
            
            [Parameter(Mandatory=$true)]
            [string]$DestinationPath
        )
        
        if (-not (Test-Path $SourceFile)) {
            Write-Warning "'$SourceFile' が見つかりません。スキップします。"
            return $false
        }
        
        Copy-Item -Path $SourceFile -Destination $DestinationPath -Force
        Write-Host "✓ Copied '$SourceFile' to '$DestinationPath'" -ForegroundColor Gray
        return $true
    }
    
    # ====================== 本処理 ======================
    
    # 現在のGitブランチを取得
    $branch = git branch --show-current

    # Gitリポジトリ内にいるか確認
    if (-not $branch) {
        Write-Error "Could not determine the current git branch.`nPlease ensure you are in a git repository."
        exit 1
    }

    Write-Host "Current branch name: $branch" -ForegroundColor Cyan

    # ローカルモード（devブランチ）の判定
    $localMode = ($branch -eq "dev")

    if ($localMode) {
        Write-Host "Running in local deployment mode..." -ForegroundColor Yellow

        & npm run lint:fix
        if ($LASTEXITCODE -ne 0) { 
            Write-Host "Linting failed. Please check the errors above." -ForegroundColor Red
            exit $LASTEXITCODE
        }

        & npm run build:dev
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Build failed. Please check the errors above." -ForegroundColor Red
            pause
            exit $LASTEXITCODE
        }

        Write-Host "Build success." -ForegroundColor Green
        Write-Host "- Run `"npm run start`" to start debugging." -ForegroundColor Green
        Write-Host "- Run `"npm run stop`" to stop debugging." -ForegroundColor Green
        Write-Host "- When debugging is complete, switch to the main branch and run this script again to deploy the release version." -ForegroundColor Green

    }
    else {
        Write-Host "Running in release branch mode..." -ForegroundColor Yellow

        # dev → main にマージ
        Write-Host "Applying latest changes from dev branch to main branch..." -ForegroundColor Cyan
        git merge dev --no-ff -m "Merge dev into main for release"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Merge failed. Please resolve conflicts and try again."
            exit $LASTEXITCODE
        }

        & npm run lint:fix
        if ($LASTEXITCODE -ne 0) { 
            Write-Host "Linting failed. Please check the errors above." -ForegroundColor Red
            exit $LASTEXITCODE
        }

        & npm run build
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Build failed. Please check the errors above." -ForegroundColor Red
            pause
            exit $LASTEXITCODE
        }

        # ==================== リリース用ファイルのコピー ====================
        Write-Host "Starting release branch deployment..." -ForegroundColor Cyan

        Copy-Contents -SourcePath "dist" `
                      -DestinationPath "..\docs" `
                      -Description "dist フォルダの中身を ..\docs にコピー"

        Copy-Contents -SourcePath "..\notes\manual" `
                      -DestinationPath "..\docs\manual" `
                      -Description "manual フォルダの中身を ..\docs\manual にコピー"

        Copy-File -SourceFile "..\README.md" `
                  -DestinationPath "..\docs"

        Write-Host "Release branch deployment complete." -ForegroundColor Green
    }
}
finally {
    # 常に元のディレクトリに戻る
    Pop-Location
}

# 5秒待機（確認用）
Start-Sleep -Seconds 5