# Dify 插件打包脚本 (Windows PowerShell)
# 用法: 在 document_cutter 的上级目录运行，或直接在本目录运行此脚本

$ErrorActionPreference = "Stop"

$PLUGIN_DIR = Split-Path -Parent $PSCommandPath
$PLUGIN_NAME = "document_cutter"
$CLI_NAME = "dify.exe"
$CLI_URL = "https://github.com/langgenius/dify-plugin-daemon/releases"

# 检查 dify CLI 是否可用
$difyCmd = Get-Command $CLI_NAME -ErrorAction SilentlyContinue
if (-not $difyCmd) {
    $difyCmd = Get-Command "dify-plugin-windows-amd64.exe" -ErrorAction SilentlyContinue
}
if (-not $difyCmd) {
    # 尝试当前目录
    if (Test-Path "$PLUGIN_DIR\dify.exe") {
        $CLI_NAME = "$PLUGIN_DIR\dify.exe"
    } elseif (Test-Path "$PLUGIN_DIR\dify-plugin-windows-amd64.exe") {
        $CLI_NAME = "$PLUGIN_DIR\dify-plugin-windows-amd64.exe"
    } else {
        Write-Host "[ERROR] dify CLI not found." -ForegroundColor Red
        Write-Host "Please download from: $CLI_URL" -ForegroundColor Yellow
        Write-Host "Then rename to dify.exe and place in PATH or this directory." -ForegroundColor Yellow
        exit 1
    }
} else {
    $CLI_NAME = $difyCmd.Source
}

Write-Host "Using CLI: $CLI_NAME" -ForegroundColor Cyan
Write-Host "Packaging plugin: $PLUGIN_DIR" -ForegroundColor Cyan

Push-Location (Split-Path -Parent $PLUGIN_DIR)
try {
    & $CLI_NAME plugin package ".\$PLUGIN_NAME"
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "[SUCCESS] Plugin packaged!" -ForegroundColor Green
        Write-Host "Generated file: $PLUGIN_NAME.difypkg" -ForegroundColor Green
        Write-Host ""
        Write-Host "Install: Dify -> Plugin Management -> Install Plugin -> Via Local File" -ForegroundColor Yellow
    } else {
        Write-Host "[FAILED] Package command returned error code $LASTEXITCODE" -ForegroundColor Red
    }
} finally {
    Pop-Location
}
