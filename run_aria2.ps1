# 检查是否通过管理员权限重新启动
$IsElevatedRestart = $args -contains "-ElevatedRestart"

# 强制使用管理员权限运行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "当前用户没有管理员权限，尝试使用管理员权限运行脚本" -ForegroundColor Yellow
    Start-Process -FilePath pwsh.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$PSCommandPath", "-ElevatedRestart" -Verb RunAs
    exit
}

# 设置控制台编码为UTF-8以支持中文显示
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$Aria2ConfPath = Join-Path $ScriptDir "aria2.conf"

$Aria2ExePath = Join-Path $ScriptDir "aria2c.exe"

# 检查aria2c.exe是否存在
if (!(Test-Path $Aria2ExePath)) {
    Write-Host "错误: 找不到 aria2c.exe 文件在路径: $Aria2ExePath" -ForegroundColor Red
    Read-Host -Prompt "Press Enter to exit..."
    exit 1
}

# 检查配置文件是否存在
if (!(Test-Path $Aria2ConfPath)) {
    Write-Host "错误: 找不到配置文件: $Aria2ConfPath" -ForegroundColor Red
    Read-Host -Prompt "Press Enter to exit..."
    exit 1
}

Write-Host "正在启动 aria2c..." -ForegroundColor Green
Write-Host "配置文件路径: $Aria2ConfPath" -ForegroundColor Cyan

try {
    # 后台启动 aria2c
    $process = Start-Process -FilePath $Aria2ExePath -ArgumentList "--conf-path=`"$Aria2ConfPath`"" -WindowStyle Hidden -PassThru
    
    if ($process) {
        Write-Host "aria2c 已在后台启动，进程ID: $($process.Id)" -ForegroundColor Green
        Write-Host "RPC 服务地址: http://localhost:6800/jsonrpc" -ForegroundColor Cyan
        
        # 如果是通过管理员权限重新启动的，直接退出
        if ($IsElevatedRestart) {
            Write-Host "管理员权限启动完成，窗口将在3秒后自动关闭..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            exit 0
        } else {
            Write-Host "aria2c 正在后台运行..." -ForegroundColor Green
        }
    } else {
        Write-Host "启动 aria2c 失败" -ForegroundColor Red
    }
} catch {
    Write-Host "启动 aria2c 时发生错误: $_" -ForegroundColor Red
    Write-Host "错误详情: $($_.Exception.Message)" -ForegroundColor Red
}

# 只有在非管理员权限重新启动时才等待用户输入
if (!$IsElevatedRestart) {
    Read-Host -Prompt "Press Enter to exit..."
}
