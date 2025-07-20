# aria2ForWin10 自动配置脚本
# 适用于 Windows 10/11 系统
# 编码：UTF-8

# 强制使用管理员权限运行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "当前用户没有管理员权限，尝试使用管理员权限运行脚本" -ForegroundColor Yellow
    Start-Process -FilePath powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$PSCommandPath" -Verb RunAs
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
        Write-Host "`n按任意键退出..." -ForegroundColor Gray
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
    Write-Host "`n1. 配置下载目录" -ForegroundColor Cyan
    
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
    Write-Host "`n2. 配置环境变量" -ForegroundColor Cyan
    
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
    Write-Host "`n3. 启动 aria2 服务" -ForegroundColor Cyan
    
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

# 函数：添加到开机自启目录
function Set-AutoStart {
    Write-Host "`n4. 配置开机自启" -ForegroundColor Cyan
    
    # 获取开机自启目录路径
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $shortcutName = "aria2启动.lnk"
    $shortcutPath = Join-Path $startupFolder $shortcutName
    $batFilePath = Join-Path $ScriptDir "run_aria2.bat"
    
    # 检查bat文件是否存在
    if (!(Test-Path $batFilePath)) {
        Write-Host "错误: 找不到启动脚本 $batFilePath" -ForegroundColor Red
        return $false
    }
    
    # 检查快捷方式是否已存在
    if (Test-Path $shortcutPath) {
        Write-Host "开机自启快捷方式已存在" -ForegroundColor Yellow
        $overwrite = Read-Host "是否要重新创建？(y/N)"
        if ($overwrite -ne 'y' -and $overwrite -ne 'Y') {
            return $true
        }
        Remove-Item $shortcutPath -Force
    }
    
    try {
        # 创建WScript.Shell对象来创建快捷方式
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $batFilePath
        $Shortcut.WorkingDirectory = $ScriptDir
        $Shortcut.Description = "aria2 开机自启服务"
        $Shortcut.Save()
        
        Write-Host "开机自启配置成功" -ForegroundColor Green
        Write-Host "快捷方式已添加到: $shortcutPath" -ForegroundColor Cyan
        return $true
    }
    catch {
        Write-Host "配置开机自启失败: $_" -ForegroundColor Red
        return $false
    }
}

# 函数：配置AriaNg快捷方式
function Set-AriaNgShortcut {
    Write-Host "`n5. 配置 AriaNg 快捷方式" -ForegroundColor Cyan
    
    # 显示选项的函数
    function Show-Options {
        param($options, $selectedIndex)
        
        # 清除控制台并重新显示
        Clear-Host
        Write-Host "找到多个浏览器，请选择:" -ForegroundColor Cyan
        Write-Host "使用 ↑↓ 方向键选择，回车确认，ESC 取消" -ForegroundColor Gray
        Write-Host ""
        
        # 显示所有选项
        for ($i = 0; $i -lt $options.Count; $i++) {
            if ($i -eq $selectedIndex) {
                Write-Host "► $($options[$i])" -ForegroundColor Black -BackgroundColor Yellow
            } else {
                Write-Host "  $($options[$i])" -ForegroundColor White
            }
        }
    }
    
    if (!(Test-Path $AriaNgLnkPath)) {
        Write-Host "AriaNg.lnk 文件不存在，将创建新的快捷方式" -ForegroundColor Yellow
    }
    
    # 查找浏览器路径
    $browserPaths = @{
        "Google Chrome" = @(
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "${env:LOCALAPPDATA}\Google\Chrome\Application\chrome.exe"
        )
        "Microsoft Edge" = @(
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
            "${env:LOCALAPPDATA}\Microsoft\Edge\Application\msedge.exe"
        )
        "Mozilla Firefox" = @(
            "${env:ProgramFiles}\Mozilla Firefox\firefox.exe",
            "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe",
            "${env:LOCALAPPDATA}\Mozilla Firefox\firefox.exe"
        )
    }
    
    # 查找可用的浏览器
    $availableBrowsers = @()
    foreach ($browserName in $browserPaths.Keys) {
        foreach ($path in $browserPaths[$browserName]) {
            if (Test-Path $path) {
                $availableBrowsers += @{
                    Name = $browserName
                    Path = $path
                }
                break
            }
        }
    }
    
    $selectedBrowserPath = $null
    
    if ($availableBrowsers.Count -eq 0) {
        Write-Host "未找到任何支持的浏览器 (Chrome, Edge, Firefox)" -ForegroundColor Yellow
        $customPath = Read-Host "请输入浏览器可执行文件的完整路径"
        
        if ($customPath -and (Test-Path $customPath)) {
            $selectedBrowserPath = $customPath
            Write-Host "使用自定义浏览器路径: $customPath" -ForegroundColor Green
        } else {
            Write-Host "无效的浏览器路径，配置失败" -ForegroundColor Red
            return $false
        }
    }
    elseif ($availableBrowsers.Count -eq 1) {
        $selectedBrowserPath = $availableBrowsers[0].Path
        Write-Host "找到浏览器: $($availableBrowsers[0].Name)" -ForegroundColor Green
    }
    else {
        # 创建选项列表
        $options = @()
        foreach ($browser in $availableBrowsers) {
            $options += "$($browser.Name) - $($browser.Path)"
        }
        $options += "自定义路径"
        
        $selectedIndex = 0
        $maxIndex = $options.Count - 1
        
        # 初始显示
        Show-Options $options $selectedIndex
        
        # 处理键盘输入
        $exitLoop = $false
        do {
            $key = [Console]::ReadKey($true)
            
            switch ($key.Key) {
                'UpArrow' {
                    if ($selectedIndex -gt 0) {
                        $selectedIndex--
                        Show-Options $options $selectedIndex
                    }
                }
                'DownArrow' {
                    if ($selectedIndex -lt $maxIndex) {
                        $selectedIndex++
                        Show-Options $options $selectedIndex
                    }
                }
                'Enter' {
                    Write-Host ""
                    if ($selectedIndex -lt $availableBrowsers.Count) {
                        $selectedBrowserPath = $availableBrowsers[$selectedIndex].Path
                        Write-Host "已选择: $($availableBrowsers[$selectedIndex].Name)" -ForegroundColor Green
                        $exitLoop = $true
                    } else {
                        # 自定义路径
                        $customPath = Read-Host "请输入浏览器可执行文件的完整路径"
                        if ($customPath -and (Test-Path $customPath)) {
                            $selectedBrowserPath = $customPath
                            Write-Host "使用自定义浏览器路径: $customPath" -ForegroundColor Green
                            $exitLoop = $true
                        } else {
                            Write-Host "无效的浏览器路径，请重新选择" -ForegroundColor Red
                            Write-Host "按任意键继续..." -ForegroundColor Gray
                            [Console]::ReadKey($true) | Out-Null
                            Show-Options $options $selectedIndex
                            continue
                        }
                    }
                }
                'Escape' {
                    Write-Host ""
                    Write-Host "已取消选择" -ForegroundColor Yellow
                    return $false
                }
            }
        } while (-not $exitLoop)
    }
    
    if (!$selectedBrowserPath) {
        Write-Host "未选择有效的浏览器，配置失败" -ForegroundColor Red
        return $false
    }
    
    try {
        # 创建或更新快捷方式
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($AriaNgLnkPath)
        $Shortcut.TargetPath = $selectedBrowserPath
        $Shortcut.Arguments = "`"$AriaNgHtmlPath`""
        $Shortcut.WorkingDirectory = Split-Path $selectedBrowserPath
        $Shortcut.Description = "AriaNg Web界面 - Aria2下载管理器"
        $Shortcut.Save()
        
        Write-Host "使用浏览器: $selectedBrowserPath" -ForegroundColor Green
        Write-Host "AriaNg 快捷方式已成功配置" -ForegroundColor Green
        Write-Host "目标: $selectedBrowserPath" -ForegroundColor Gray
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
    
    Write-Host "`n6. 固定到开始屏幕" -ForegroundColor Cyan
    
    if (!(Test-Path $ShortcutPath)) {
        Write-Host "快捷方式文件不存在，无法固定到开始屏幕" -ForegroundColor Red
        return $false
    }
    
    try {
        # 首先尝试复制快捷方式到开始菜单程序文件夹
        $startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
        $targetPath = Join-Path $startMenuPath "AriaNg.lnk"
        
        # 确保目标目录存在
        if (!(Test-Path $startMenuPath)) {
            New-Item -Path $startMenuPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item $ShortcutPath $targetPath -Force
        Write-Host "快捷方式已复制到开始菜单程序文件夹" -ForegroundColor Green
        Write-Host "路径: $targetPath" -ForegroundColor Gray
        
        # 等待文件系统同步
        Start-Sleep -Milliseconds 500
        
        # 验证文件是否成功复制
        if (!(Test-Path $targetPath)) {
            throw "快捷方式复制失败"
        }
        
        # 尝试使用Shell.Application固定到开始屏幕
        try {
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($startMenuPath)
            $item = $folder.ParseName("AriaNg.lnk")
            
            if ($item) {
                # 查找固定到开始屏幕的动词
                $pinVerb = $null
                foreach ($verb in $item.Verbs()) {
                    # 支持中文和英文系统，扩展匹配模式
                    if ($verb.Name -match "固定.*开始.*屏幕" -or 
                        $verb.Name -match "Pin.*Start" -or 
                        $verb.Name -match "固定到.*开始" -or 
                        $verb.Name -eq '固定到"开始"屏幕(&P)' -or
                        $verb.Name -match "Pin to Start" -or
                        $verb.Name -match "固定到开始菜单") {
                        $pinVerb = $verb
                        break
                    }
                }
                
                if ($pinVerb) {
                    # 尝试以管理员权限执行
                    try {
                        $pinVerb.DoIt()
                        # 等待操作完成
                        Start-Sleep -Milliseconds 1000
                        Write-Host "AriaNg 快捷方式已成功固定到开始屏幕" -ForegroundColor Green
                        Write-Host "您可以在开始菜单中找到 AriaNg 磁贴" -ForegroundColor Gray
                        return $true
                    }
                    catch {
                        # 如果是权限问题，尝试替代方案
                        if ($_.Exception.HResult -eq 0x80070005 -or $_.Exception.Message -match "拒绝访问") {
                            Write-Host "系统安全策略限制了自动固定功能" -ForegroundColor Yellow
                            Write-Host "这是Windows 10/11的正常安全限制" -ForegroundColor Gray
                        } else {
                            Write-Host "固定操作失败: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                        throw $_
                    }
                } else {
                    Write-Host "未找到固定到开始屏幕的选项" -ForegroundColor Yellow
                    # 列出所有可用的动词以便调试
                    Write-Host "可用操作:" -ForegroundColor Gray
                    foreach ($verb in $item.Verbs()) {
                        Write-Host "  - $($verb.Name)" -ForegroundColor Gray
                    }
                    throw "无固定选项"
                }
            } else {
                throw "无法在开始菜单中找到快捷方式"
            }
        }
        catch {
            # Shell.Application方法失败，提供手动操作指导
            Write-Host "自动固定失败，请手动完成以下步骤:" -ForegroundColor Yellow
            Write-Host "1. 按 Win 键打开开始菜单" -ForegroundColor White
            Write-Host "2. 在程序列表中找到 'AriaNg'" -ForegroundColor White
            Write-Host "3. 右键点击 'AriaNg' -> 选择 '固定到开始屏幕'" -ForegroundColor White
            Write-Host "4. 或者在文件资源管理器中打开: $targetPath" -ForegroundColor White
            Write-Host "5. 右键点击快捷方式 -> 选择 '固定到开始屏幕'" -ForegroundColor White
            return $true
        }
    }
    catch {
        Write-Host "操作失败: $_" -ForegroundColor Red
        
        # 最后的备用方案：尝试复制到公共开始菜单
        try {
            $publicStartMenuPath = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs"
            $publicTargetPath = Join-Path $publicStartMenuPath "AriaNg.lnk"
            
            if (Test-Path $publicStartMenuPath) {
                Copy-Item $ShortcutPath $publicTargetPath -Force
                Write-Host "快捷方式已复制到公共开始菜单" -ForegroundColor Green
                Write-Host "路径: $publicTargetPath" -ForegroundColor Gray
                Write-Host "请手动固定到开始屏幕" -ForegroundColor Yellow
                return $true
            }
        }
        catch {
            Write-Host "所有自动操作都失败了" -ForegroundColor Red
            Write-Host "请手动将快捷方式 $ShortcutPath 复制到开始菜单" -ForegroundColor Yellow
        }
        
        return $false
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
    Write-Host "`n=== 配置完成摘要 ===" -ForegroundColor Green
    $successCount = ($results | Where-Object { $_ -eq $true }).Count
    $totalCount = $results.Count
    Write-Host "成功: $successCount/$totalCount 项配置" -ForegroundColor Green
    
    if ($successCount -eq $totalCount) {
        Write-Host "`n所有配置项都已成功完成！" -ForegroundColor Green
        Write-Host "现在可以使用 AriaNg.lnk 启动 Web 界面进行下载管理" -ForegroundColor Cyan
    }
    else {
        Write-Host "`n部分配置项未成功，请检查上述错误信息" -ForegroundColor Yellow
    }
    
    Write-Host "`n按任意键退出..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# 运行主函数
Main