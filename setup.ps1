# aria2ForWin10 自动配置脚本
# 适用于 Windows 10/11 系统
# 编码：UTF-8

# 强制使用管理员权限运行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "当前用户没有管理员权限，尝试使用管理员权限运行脚本" -ForegroundColor Yellow
    Start-Process -FilePath powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File \"$PSCommandPath\"" -Verb RunAs
    exit
}

# 检查并安装 PowerShell 7.x
function Ensure-PowerShell7 {
    Write-Host "=== 检查 PowerShell 版本 ===" -ForegroundColor Green
    
    # 检查当前 PowerShell 版本
    $currentVersion = $PSVersionTable.PSVersion
    Write-Host "当前 PowerShell 版本: $currentVersion" -ForegroundColor Yellow
    
    # 如果已经是 PowerShell 7.x，直接返回
    if ($currentVersion.Major -ge 7) {
        Write-Host "已使用 PowerShell 7.x，继续执行配置" -ForegroundColor Green
        return $true
    }
    
    Write-Host "检测到 Windows PowerShell 5.x，正在安装 PowerShell 7.x..." -ForegroundColor Yellow
    
    # 检查 winget 是否可用
    try {
        $null = Get-Command winget -ErrorAction Stop
        Write-Host "找到 winget，开始安装 PowerShell 7.x" -ForegroundColor Green
    }
    catch {
        Write-Host "未找到 winget，请手动安装 PowerShell 7.x" -ForegroundColor Red
        Write-Host "下载地址: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
        return $false
    }
    
    # 使用 winget 安装 PowerShell 7.x
    try {
        Write-Host "正在安装 PowerShell 7.x，请稍候..." -ForegroundColor Yellow
        $installResult = winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PowerShell 7.x 安装成功！" -ForegroundColor Green
        }
        else {
            Write-Host "PowerShell 7.x 可能已安装或安装遇到问题" -ForegroundColor Yellow
        }
        
        # 检查 pwsh 是否可用
        $pwshPath = $null
        $possiblePaths = @(
            "$env:ProgramFiles\PowerShell\7\pwsh.exe",
            "$env:ProgramFiles(x86)\PowerShell\7\pwsh.exe",
            "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $pwshPath = $path
                break
            }
        }
        
        if (!$pwshPath) {
            # 尝试从环境变量中找到 pwsh
            try {
                $pwshPath = (Get-Command pwsh -ErrorAction Stop).Source
            }
            catch {
                Write-Host "无法找到 pwsh.exe，请重启计算机后再试" -ForegroundColor Red
                return $false
            }
        }
        
        Write-Host "找到 PowerShell 7.x: $pwshPath" -ForegroundColor Green
        Write-Host "正在使用 PowerShell 7.x 重新执行脚本..." -ForegroundColor Cyan
        
        # 使用 PowerShell 7.x 重新执行当前脚本
        $currentScript = $MyInvocation.MyCommand.Definition
        & "$pwshPath" -File "$currentScript" -SkipPowerShellCheck
        
        # 退出当前 PowerShell 5.x 进程
        exit $LASTEXITCODE
    }
    catch {
        Write-Host "安装 PowerShell 7.x 失败: $_" -ForegroundColor Red
        Write-Host "请手动安装 PowerShell 7.x 后重新运行脚本" -ForegroundColor Yellow
        return $false
    }
}

# 检查是否跳过 PowerShell 检查（用于避免无限递归）
if (-not $args -contains "-SkipPowerShellCheck") {
    if (-not (Ensure-PowerShell7)) {
        Write-Host "\n按任意键退出..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

# 设置控制台编码为UTF-8以支持中文显示
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 获取脚本所在目录
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Aria2ConfPath = Join-Path $ScriptDir "aria2.conf"
$AriaNgHtmlPath = Join-Path $ScriptDir "AriaNg.html"
$AriaNgLnkPath = Join-Path $ScriptDir "AriaNg.lnk"

Write-Host "=== aria2forwin10 自动配置脚本 ===" -ForegroundColor Green
Write-Host "脚本目录: $ScriptDir" -ForegroundColor Yellow

# 函数：检查管理员权限
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 函数：配置下载目录
function Set-DownloadDirectory {
    Write-Host "\n1. 配置下载目录" -ForegroundColor Cyan
    
    # 读取当前配置
    if (Test-Path $Aria2ConfPath) {
        $content = Get-Content $Aria2ConfPath -Encoding UTF8
        $currentDir = ($content | Where-Object { $_ -match '^dir=' }) -replace '^dir=', ''
        Write-Host "当前下载目录: $currentDir" -ForegroundColor Yellow
    }
    
    # 询问用户是否要修改
    $newDir = Read-Host "请输入新的下载目录路径（直接回车保持当前设置）"
    
    if ($newDir -and $newDir.Trim() -ne "") {
        # 创建目录（如果不存在）
        if (!(Test-Path $newDir)) {
            try {
                New-Item -ItemType Directory -Path $newDir -Force | Out-Null
                Write-Host "已创建下载目录: $newDir" -ForegroundColor Green
            }
            catch {
                Write-Host "创建目录失败: $_" -ForegroundColor Red
                return $false
            }
        }
        
        # 更新配置文件
        try {
            $content = Get-Content $Aria2ConfPath -Encoding UTF8
            $content = $content -replace '^dir=.*', "dir=$newDir"
            $content | Set-Content $Aria2ConfPath -Encoding UTF8
            Write-Host "下载目录已更新为: $newDir" -ForegroundColor Green
        }
        catch {
            Write-Host "更新配置文件失败: $_" -ForegroundColor Red
            return $false
        }
    }
    return $true
}

# 函数：配置环境变量
function Set-EnvironmentPath {
    Write-Host "\n2. 配置环境变量" -ForegroundColor Cyan
    
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
    
    if ($currentPath -notlike "*$ScriptDir*") {
        try {
            $newPath = $currentPath + ";" + $ScriptDir
            [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
            Write-Host "已将项目目录添加到用户环境变量" -ForegroundColor Green
            Write-Host "注意: 需要重新打开命令行窗口才能生效" -ForegroundColor Yellow
        }
        catch {
            Write-Host "添加环境变量失败: $_" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "项目目录已在环境变量中" -ForegroundColor Yellow
    }
    return $true
}

# 函数：启动aria2
function Start-Aria2Service {
    Write-Host "\n3. 启动 aria2 服务" -ForegroundColor Cyan
    
    # 检查是否已经运行
    $aria2Process = Get-Process -Name "aria2c" -ErrorAction SilentlyContinue
    if ($aria2Process) {
        Write-Host "aria2 服务已在运行中 (PID: $($aria2Process.Id))" -ForegroundColor Yellow
        return $true
    }
    
    try {
        $aria2Exe = Join-Path $ScriptDir "aria2c.exe"
        Start-Process -FilePath $aria2Exe -ArgumentList "--conf-path=`"$Aria2ConfPath`"" -WindowStyle Hidden
        Start-Sleep -Seconds 2
        
        # 验证启动
        $aria2Process = Get-Process -Name "aria2c" -ErrorAction SilentlyContinue
        if ($aria2Process) {
            Write-Host "aria2 服务启动成功 (PID: $($aria2Process.Id))" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "aria2 服务启动失败" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "启动 aria2 失败: $_" -ForegroundColor Red
        return $false
    }
}

# 函数：创建开机自启任务
function Set-AutoStart {
    Write-Host "\n4. 配置开机自启" -ForegroundColor Cyan
    
    if (!(Test-Administrator)) {
        Write-Host "配置开机自启需要管理员权限，跳过此步骤" -ForegroundColor Yellow
        Write-Host "请以管理员身份运行脚本来配置开机自启" -ForegroundColor Yellow
        return $false
    }
    
    $taskName = "aria2启动"
    
    # 检查任务是否已存在
    try {
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Host "开机自启任务已存在" -ForegroundColor Yellow
            $overwrite = Read-Host "是否要重新创建？(y/N)"
            if ($overwrite -ne 'y' -and $overwrite -ne 'Y') {
                return $true
            }
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }
    }
    catch {
        # 任务不存在，继续创建
    }
    
    try {
        # 创建任务动作
        $aria2Exe = Join-Path $ScriptDir "aria2c.exe"
        $action = New-ScheduledTaskAction -Execute $aria2Exe -Argument "--conf-path=`"$Aria2ConfPath`""
        
        # 创建触发器（登录时）
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        
        # 创建任务设置
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        
        # 创建任务主体（以当前用户身份运行）
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
        
        # 注册任务
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "aria2 开机自启服务" | Out-Null
        
        Write-Host "开机自启任务创建成功" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "创建开机自启任务失败: $_" -ForegroundColor Red
        return $false
    }
}

# 函数：配置AriaNg快捷方式
function Set-AriaNgShortcut {
    Write-Host "\n5. 配置 AriaNg 快捷方式" -ForegroundColor Cyan
    
    if (!(Test-Path $AriaNgLnkPath)) {
        Write-Host "AriaNg.lnk 文件不存在，将创建新的快捷方式" -ForegroundColor Yellow
    }
    
    # 查找Chrome路径
    $chromePaths = @(
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "${env:LOCALAPPDATA}\Google\Chrome\Application\chrome.exe"
    )
    
    $chromePath = $null
    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            $chromePath = $path
            break
        }
    }
    
    if (!$chromePath) {
        Write-Host "未找到 Chrome 浏览器，请手动安装 Chrome 或配置快捷方式" -ForegroundColor Red
        return $false
    }
    
    try {
        # 创建或更新快捷方式
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($AriaNgLnkPath)
        $Shortcut.TargetPath = $chromePath
        $Shortcut.Arguments = "`"$AriaNgHtmlPath`""
        $Shortcut.WorkingDirectory = $ScriptDir
        $Shortcut.Description = "AriaNg Web UI for aria2"
        $Shortcut.Save()
        
        Write-Host "找到 Chrome 浏览器: $chromePath" -ForegroundColor Green
        Write-Host "AriaNg 快捷方式已成功配置" -ForegroundColor Green
        Write-Host "目标: $chromePath" -ForegroundColor Gray
        Write-Host "参数: `"$AriaNgHtmlPath`"" -ForegroundColor Gray
        
        # 尝试固定到开始屏幕
        Set-StartMenuTile -ShortcutPath $AriaNgLnkPath
        
        return $true
    }
    catch {
        Write-Host "配置 AriaNg 快捷方式失败: $_" -ForegroundColor Red
        return $false
    }
}

# 函数：固定快捷方式到开始屏幕
function Set-StartMenuTile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ShortcutPath
    )
    
    Write-Host "\n6. 固定到开始屏幕" -ForegroundColor Cyan
    
    if (!(Test-Path $ShortcutPath)) {
        Write-Host "快捷方式文件不存在，无法固定到开始屏幕" -ForegroundColor Red
        return $false
    }
    
    try {
        # 方法1：使用Shell.Application COM对象
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Split-Path $ShortcutPath))
        $item = $folder.ParseName((Split-Path $ShortcutPath -Leaf))
        
        # 获取上下文菜单项
        $verbs = $item.Verbs()
        $pinVerb = $null
        
        # 查找"固定到开始屏幕"动作（支持中英文）
        foreach ($verb in $verbs) {
            if ($verb.Name -match "固定到.*开始" -or $verb.Name -match "Pin to Start") {
                $pinVerb = $verb
                break
            }
        }
        
        if ($pinVerb) {
            $pinVerb.DoIt()
            Write-Host "AriaNg 快捷方式已成功固定到开始屏幕" -ForegroundColor Green
            Write-Host "您可以在开始菜单中找到 AriaNg 磁贴" -ForegroundColor Gray
            return $true
        } else {
            # 方法2：使用PowerShell的Start-Process调用explorer
            Write-Host "尝试使用备用方法固定到开始屏幕..." -ForegroundColor Yellow
            
            # 复制快捷方式到开始菜单程序文件夹
            $startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
            $targetPath = Join-Path $startMenuPath "AriaNg.lnk"
            
            Copy-Item $ShortcutPath $targetPath -Force
            Write-Host "快捷方式已复制到开始菜单程序文件夹" -ForegroundColor Green
            Write-Host "路径: $targetPath" -ForegroundColor Gray
            
            # 提示用户手动固定
            Write-Host "\n请手动完成以下步骤来固定磁贴:" -ForegroundColor Yellow
            Write-Host "1. 按 Win 键打开开始菜单" -ForegroundColor White
            Write-Host "2. 在程序列表中找到 'AriaNg'" -ForegroundColor White
            Write-Host "3. 右键点击 'AriaNg' -> 选择 '固定到开始屏幕'" -ForegroundColor White
            
            return $true
        }
    }
    catch {
        Write-Host "固定到开始屏幕时出现错误: $_" -ForegroundColor Red
        
        # 备用方案：复制到开始菜单
        try {
            $startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
            $targetPath = Join-Path $startMenuPath "AriaNg.lnk"
            
            Copy-Item $ShortcutPath $targetPath -Force
            Write-Host "快捷方式已复制到开始菜单程序文件夹" -ForegroundColor Green
            Write-Host "您可以在开始菜单中找到并手动固定 AriaNg" -ForegroundColor Gray
            
            return $true
        }
        catch {
            Write-Host "复制到开始菜单也失败: $_" -ForegroundColor Red
            return $false
        }
    }
}

# 主函数
function Main {
    Write-Host "检查管理员权限..." -ForegroundColor Yellow
    if (Test-Administrator) {
        Write-Host "当前以管理员身份运行" -ForegroundColor Green
    }
    else {
        Write-Host "当前以普通用户身份运行（部分功能可能受限）" -ForegroundColor Yellow
    }
    
    # 执行配置步骤
    $results = @()
    $results += Set-DownloadDirectory
    $results += Set-EnvironmentPath
    $results += Start-Aria2Service
    $results += Set-AutoStart
    $results += Set-AriaNgShortcut
    
    # 显示结果摘要
    Write-Host "\n=== 配置完成摘要 ===" -ForegroundColor Green
    $successCount = ($results | Where-Object { $_ -eq $true }).Count
    $totalCount = $results.Count
    Write-Host "成功: $successCount/$totalCount 项配置" -ForegroundColor Green
    
    if ($successCount -eq $totalCount) {
        Write-Host "\n所有配置项都已成功完成！" -ForegroundColor Green
        Write-Host "现在可以使用 AriaNg.lnk 启动 Web 界面进行下载管理" -ForegroundColor Cyan
    }
    else {
        Write-Host "\n部分配置项未成功，请检查上述错误信息" -ForegroundColor Yellow
    }
    
    Write-Host "\n按任意键退出..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# 运行主函数
Main