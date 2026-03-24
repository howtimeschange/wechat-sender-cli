# 微信批量发送助手 CLI

跨平台命令行工具，支持 macOS 和 Windows，可执行微信批量发送和定时发送。

## 功能特性

- **批量发送** — 读取 Excel 模板，自动向多个联系人/群聊发送文字/图片/文字+图片
- **定时发送** — 支持指定发送时间，到点自动发送
- **重复发送** — 支持 daily / weekly / workday 重复任务
- **速率控制** — 可配置发送间隔和每分钟最大条数，防封号
- **状态回写** — 发送结果实时写入 Excel，无需人工核对
- **风控保护** — 关键词拦截 / 限速 / 熔断 / 白名单
- **跨平台** — macOS（AppleScript）/ Windows（uiautomation）

## 系统要求

| 平台 | 系统版本 | Python | 自动化 |
|------|----------|--------|--------|
| macOS | macOS 12+ | 3.10+ | ✅ AppleScript |
| Windows | Windows 10+ | 3.10+ | ✅ uiautomation |
| Linux | 任意 | 3.10+ | ❌ 仅查看/配置 |

## 一键安装

### macOS

```bash
bash <(curl -fsSL https://raw.githubusercontent.com/howtimeschange/wechat-sender-cli/main/install.sh)
```

### Windows（管理员 PowerShell）

```powershell
irm https://raw.githubusercontent.com/howtimeschange/wechat-sender-cli/main/install.ps1 | iex
```

手动安装：

```bash
git clone https://github.com/howtimeschange/wechat-sender-cli.git
cd wechat-sender-cli
bash install.sh      # macOS
# 或
install.bat          # Windows
```

## 快速开始

```bash
# 1. 首次配置（交互向导，会提示输入表格路径等）
python3 app/cli.py setup

# 2. 查看当前配置
python3 app/cli.py config

# 3. 查看任务状态
python3 app/cli.py status

# 4. 立即发送
python3 app/cli.py send

# 5. 守护进程模式（持续监控表格，有新任务自动发送）
python3 app/cli.py daemon
```

## 配置项

```bash
# 查看所有配置
python3 app/cli.py config

# 修改发送间隔（秒）
python3 app/cli.py config send_interval 5

# 修改每分钟最大条数
python3 app/cli.py config max_per_minute 8

# 开启模拟运行（不真实发送）
python3 app/cli.py config dry_run true

# 修改轮询间隔（守护进程模式）
python3 app/cli.py config poll_seconds 15

# 修改表格路径
python3 app/cli.py config excel_path /Users/xx/Downloads/tasks.xlsx
```

## Excel 模板

在 Excel 中新建表格，Sheet 名称固定为「发送任务」，结构如下：

| # | * 应用 | * 联系人/群聊 | * 消息类型 | * 文字内容 | 图片路径 | 发送时间 | 重复 | 备注 | 状态 |
|---|--------|--------------|-----------|-----------|----------|----------|------|------|------|
| 1 | 微信 | 张三 | 文字 | 你好，这是一条测试消息 | | | | | 发送成功 |
| 2 | 微信 | 产品讨论群 | 文字+图片 | 活动图片来了 | /Users/xx/pic.png | 2025-01-01 09:00 | | | 待发送 |

**列说明：**

| 列名 | 是否必填 | 说明 | 示例 |
|------|---------|------|------|
| * 应用 | ✅ | 固定填「微信」 | 微信 |
| * 联系人/群聊 | ✅ | 精确的微信昵称或群名（用于搜索） | 张三 |
| * 消息类型 | ✅ | 文字 / 图片 / 文字+图片 | 文字 |
| * 文字内容 | ✅ | 要发送的文本 | 你好！ |
| 图片路径 | 图片时必填 | 本地图片绝对路径 | /Users/xx/a.png |
| 发送时间 | 可选 | 格式 `2025-01-01 14:30`，留空=立即发送 | 2025-01-01 14:30 |
| 重复 | 可选 | daily / weekly / workday / 留空=不重复 | daily |
| 备注 | 可选 | 任意备注，不影响发送 | 测试用 |
| 状态 | 自动 | 发送后自动填充，勿手动编辑 | 发送成功 |

## 风控说明

内置以下保护机制，可有效降低微信封号风险：

1. **关键词拦截** — 检测到转账、借钱、合同、发票等敏感词时自动跳过
2. **限速** — 同联系人每分钟最多 N 条（可配置，默认 8 条）
3. **熔断** — 连续失败 3 次自动暂停 30 秒
4. **白名单** — 可设置仅对指定联系人启用自动化

## Windows 自动回复模式

Windows 版额外支持监控自动回复功能：

```cmd
python scripts/watch_and_reply_win.py --contact 文件传输助手 --poll 2 --dry
```

参数说明：

| 参数 | 说明 | 默认值 |
|------|------|--------|
| --contact | 监控的联系人 | 文件传输助手 |
| --poll | 轮询间隔（秒） | 2 |
| --dry | 模拟运行，不真实发送 | false |

## 权限说明

**macOS：**
需在「系统设置 → 隐私与安全性 → 辅助功能」中授予 Terminal 或 Python 权限，否则 AppleScript 自动化会被系统拦截。

**Windows：**
首次运行需要以管理员身份启动，并允许 UI 自动化访问。

## 项目结构

```
wechat-sender-cli/
├── app/
│   └── cli.py                 # 跨平台 CLI 主程序
├── scripts/
│   ├── wechat_send_mac.applescript   # macOS 自动化核心
│   ├── wechat_send_win.py     # Windows 自动化核心
│   └── watch_and_reply_win.py # Windows 自动回复
├── install.sh                 # macOS 安装脚本
├── install.bat               # Windows cmd 安装脚本
├── install.ps1               # Windows PowerShell 一键脚本
├── requirements.txt          # 基础依赖
├── requirements-win.txt       # Windows 自动化依赖
├── config.yaml               # 配置文件（自动生成）
└── README.md
```

## 常见问题

**Q: 提示 "osascript 不允许发送按键" 或 error -1002？**
A: macOS 辅助功能权限未授权。请到「系统设置 → 隐私与安全性 → 辅助功能」勾选 Terminal 或 Python。

**Q: Windows 提示找不到微信窗口？**
A: 请先打开 PC 微信并登录，确认窗口标题为「微信」。

**Q: 发送失败，显示 "不支持的消息类型"？**
A: 消息类型列只能填「文字」「图片」「文字+图片」，注意不要有多余空格。

**Q: 图片发送失败？**
A: Windows 版请使用绝对路径（如 `C:\Users\xx\pic.png`），图片格式支持 PNG/JPG/BMP。

**Q: 被封号了怎么办？**
A: 降低发送频率（增加间隔、减少每分钟条数），避免短时间内向大量陌生人发送相同内容。
