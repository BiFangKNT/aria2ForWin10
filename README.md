# aria2ForWin10

一个 aria2 的 Windows 10 Web UI 整合包。本项目整合了原版 aria2 的 win-64位 发行版和 AriaNg 的单文件发行版，为用户提供开箱即用的解决方案。

我会不定期同步上游更新，欢迎 PR。

## 目录

- [项目组成](#项目组成)
- [使用方法](#使用方法)
  - [快速开始（推荐）](#快速开始推荐)
  - [手动配置方法](#手动配置方法)
  - [AriaNg 使用方法](#ariang-使用方法)
  - [开始屏幕磁贴](#开始屏幕磁贴)
- [推荐浏览器扩展](#推荐浏览器扩展)
- [配置说明](#配置说明)
- [注意事项](#注意事项)
- [故障排除](#故障排除)
- [许可证](#许可证)
- [相关链接](#相关链接)

## 项目组成

- **aria2c.exe**: aria2 官方 Windows 64位版本
- **aria2.conf**: aria2 配置文件
- **AriaNg.html**: AriaNg 单文件版本，为 aria2 的 Web UI
- **AriaNg.lnk**: AriaNg 快捷方式

## 使用方法

### 快速开始（推荐）
1. 双击运行 `一键启动.bat` 文件
2. 批处理文件会自动请求管理员权限并启动 PowerShell 脚本
3. 脚本会自动检测并安装 PowerShell 7.x（如果需要）
4. 按照脚本提示完成配置
5. 脚本会自动完成下载目录配置、环境变量设置、aria2启动和开机自启等所有步骤

> **脚本特性**：
> - **自动版本管理**：脚本会自动检测当前 PowerShell 版本，如果是 Windows PowerShell 5.x，会使用 winget 自动安装 PowerShell 7.x 并切换执行环境
> - **兼容性保证**：确保所有功能在最新的 PowerShell 环境中正常运行
> - **无需手动干预**：整个过程完全自动化
> - **开始屏幕集成**：自动将 AriaNg 快捷方式固定到开始屏幕磁贴
> - **管理员权限处理**：批处理文件会自动请求管理员权限，确保所有功能正常运行
>
> **注意事项**：
> - **推荐使用批处理文件**：`一键启动.bat` 会自动处理权限和执行策略问题
> - 如果直接运行 PowerShell 脚本遇到执行策略限制，请以管理员身份打开 PowerShell 并执行：
>   ```powershell
>   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
>   ```
> - 如果系统没有 winget，脚本会提示手动安装 PowerShell 7.x

### 手动配置方法

如果你更喜欢手动配置或脚本运行遇到问题，可以按照以下步骤操作：

#### 使用准备

0. **下载解压**
   - 下载本仓库的压缩包
   - 解压到你喜欢的路径

1. **配置下载目录**
   - 编辑 `aria2.conf` 文件
   - 修改第4行 `dir=D:\ariaDownload` 为你喜欢的下载位置

2. **配置 AriaNg 快捷方式**
   - 右键 `AriaNg.lnk` 打开属性
   - "目标"项里面有两个路径：
     - 前者为浏览器位置，预设为 Chrome 默认位置
     - 后者按需改为项目内 `AriaNg.html` 的绝对路径
   - 应用后右键该文件，固定到"开始"屏幕

3. **添加环境变量**
   - 将项目根目录添加进系统环境变量 PATH 中

#### aria2 首次启动方法

设 `$aria2Conf` 为 `aria2.conf` 的路径

管理员权限打开 PowerShell，执行以下命令：

```powershell
Start-Process aria2c -ArgumentList "--conf-path=$aria2Conf" -WindowStyle Hidden
```

#### aria2 开机自启设置方法

利用任务计划程序设置开机自动启动：

1. 按 `Win + R` 输入 `taskschd.msc` 打开任务计划程序
2. 选择"创建任务"
3. 在"常规"页面：
   - 名称填写"aria2启动"
   - 勾选"使用最高权限运行"
4. 在"触发器"标签页：
   - 新建触发器，选"登录时"
5. 在"操作"标签页：
   - 新建操作，操作选择"启动程序"
   - 程序或脚本填写 `aria2c` 的路径（如果已经添加了环境变量，可直接写 `aria2c`）
   - 添加参数填写你的启动参数：
     ```
     --conf-path=$aria2Conf
     ```
6. 确定保存

#### AriaNg 使用方法

1. 双击打开 AriaNg 快捷方式，或“开始”屏幕点击 AriaNg 磁贴
2. 如 aria2 已经启动，将会自动连接
3. 即可开始使用 AriaNg

##### 推荐配置

###### aria2设置

- 基本设置 -> 断点续传 ：是
- HTTP/FTP/SFTP 设置 -> 最小文件分片大小 ：1MB
- BitTorrent 设置 -> 
    - BT 服务器地址：从 `https://cf.trackerslist.com/best_aria2.txt` 中复制粘贴
    - 全局最大上传速度：2M
- 高级设置 -> 
    - 始终断点续传：是
    - 优化并发下载：true

#### 开始屏幕磁贴

脚本会自动尝试将 AriaNg 快捷方式固定到 Windows 开始屏幕，方便快速访问。

##### 自动固定功能
- 脚本会尝试使用 Windows Shell API 自动固定磁贴
- 如果自动固定失败，会将快捷方式复制到开始菜单程序文件夹
- 支持中英文 Windows 系统

##### 手动固定步骤
如果自动固定失败，请按以下步骤手动操作：
1. 按 `Win` 键打开开始菜单
2. 在程序列表中找到 "AriaNg"
3. 右键点击 "AriaNg" → 选择 "固定到开始屏幕"
4. AriaNg 磁贴将出现在开始屏幕上

## 推荐浏览器扩展

### 猫抓 (CatCatch)
一个强大的资源嗅探浏览器扩展，可以自动检测网页中的媒体资源。

**安装方式**：
- Chrome/Edge: 在扩展商店搜索"猫抓"或"CatCatch"
   - [Chrome 扩展商店](https://chrome.google.com/webstore/detail/jfedfbgedapdagkghmgibemcoggfppbb)
   - [Edge 扩展商店](https://microsoftedge.microsoft.com/addons/detail/%E7%8C%AB%E6%8A%93/oohmdefbjalncfplafanlagojlakmjci)
- Firefox: 在附加组件商店搜索"猫抓"
   - [Firefox 附加组件商店](https://addons.mozilla.org/addon/cat-catch/)

**配置建议**：
1. 安装后点击扩展图标进入设置
2. 在设置中配置：
   - 后缀 -> mp4 -> 过滤大小 -> 1024
   - Aria2 RPC -> 启用
   - M3U8解析器 ->
      - 启用
      - 下载线程 -> 8~32
      - mp4格式 -> 启用

## 配置说明

### aria2.conf 主要配置项

```ini
enable-rpc=true          # 启用 RPC 服务
rpc-listen-all=true      # RPC 监听所有地址
rpc-allow-origin-all=true # 允许所有来源的 RPC 请求
dir=D:\ariaDownload      # 下载目录（请修改为你的路径）
```

## 注意事项

- 首次使用需要管理员权限启动 aria2
- 确保防火墙允许 aria2 通过
- 建议定期更新 aria2 和 AriaNg 到最新版本
- 下载目录需要有足够的磁盘空间

## 故障排除

### 常见问题

1. **aria2 启动失败**
   - 检查配置文件路径是否正确
   - 确认是否有管理员权限
   - 检查端口是否被占用

2. **AriaNg 无法连接**
   - 确认 aria2 RPC 服务已启动
   - 检查防火墙设置
   - 验证 RPC 端口配置

3. **下载失败**
   - 检查下载目录是否存在且有写入权限
   - 确认网络连接正常
   - 检查 URL 是否有效

## 许可证

本项目为整合包，各组件遵循其原有许可证：
- aria2: [GPLv2+](https://github.com/aria2/aria2/blob/master/COPYING)
- AriaNg: [MIT License](https://github.com/mayswind/AriaNg/blob/master/LICENSE)

## 相关链接

- [aria2 官方项目](https://github.com/aria2/aria2)
- [AriaNg 官方项目](https://github.com/mayswind/AriaNg)
- [aria2 官方文档](https://aria2.github.io/manual/en/html/)