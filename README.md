# 微信批量发送助手 CLI

跨平台命令行工具，支持 macOS 和 Windows，可执行微信批量发送和自动回复。

## 系统要求

- **macOS**: Python 3.10+，支持微信自动化发送（通过 AppleScript）
- **Windows**: Python 3.10+，支持微信自动化发送（通过 uiautomation）
- **Linux**: Python 3.10+，仅支持查看/配置

## 快速安装

### macOS（一键安装）

```bash
bash <(curl -fsSL https://raw.githubusercontent.com/howtimeschange/wechat-sender-cli/main/install.sh)
```

或手动：

```bash
git clone https://github.com/howtimeschange/wechat-sender-cli.git
cd wechat-sender-cli
bash install.sh
```

### Windows

```cmd
git clone https://github.com/howtimeschange/wechat-sender-cli.git
cd wechat-sender-cli
install.bat
```

## 依赖

**所有平台：**
- openpyxl
- PyYAML
- rich

**macOS：**
- 无额外依赖（使用系统 osascript）

**Windows 自动化（可选）：**
- uiautomation
- pyperclip
- pywin32
- Pillow

Windows 安装脚本会自动安装以上依赖。

## 使用方法

```bash
# 首次配置（交互向导）
python3 app/cli.py setup

# 查看当前配置
python3 app/cli.py config

# 查看任务状态
python3 app/cli.py status

# 立即发送待处理任务（macOS / Windows）
python3 app/cli.py send

# 守护进程模式（后台轮询，macOS / Windows）
python3 app/cli.py daemon

# 修改配置项
python3 app/cli.py config send_interval 5
python3 app/cli.py config max_per_minute 8

# 查看模板格式
python3 app/cli.py template
```

## 配置说明

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| excel_path | Excel 表格文件路径 | - |
| poll_seconds | 守护进程轮询间隔（秒） | 15 |
| send_interval | 两条消息间隔（秒） | 5 |
| max_per_minute | 每分钟最大条数 | 8 |
| dry_run | 模拟运行（不真实发送） | false |

## Excel 表格模板

| 列名 | 说明 | 示例 |
|------|------|------|
| * 应用 | 固定填 微信 | 微信 |
| * 联系人/群聊 | 对方微信名或群名 | 张三 |
| * 消息类型 | 文字/图片/文字+图片 | 文字 |
| * 文字内容 | 要发送的文本 | 你好！ |
| 图片路径 | 本地图片绝对路径 | /Users/xx/pic.png |
| 发送时间 | 格式 2025-01-01 14:30 | 2025-01-01 14:30 |
| 重复 | daily/weekly/workday/空 | daily |
| 状态 | 自动填充，勿手动编辑 | 发送成功 |

## 风控说明

内置以下保护机制：

1. **关键词拦截** — 自动跳过包含以下词的发送指令：转账、借钱、合同、发票、付款、汇款、账户、银行卡、密码
2. **限速** — 同联系人每分钟最多 3 条（可配置）
3. **熔断** — 连续失败 3 次自动暂停 30 秒
4. **白名单** — 可设置只对指定联系人启用自动化

## Windows 自动回复（watch_and_reply_win.py）

```cmd
python scripts/watch_and_reply_win.py --contact 文件传输助手 --poll 2 --dry
```

参数：
- `--contact` 监控的联系人（默认：文件传输助手）
- `--poll` 轮询间隔秒数（默认：2）
- `--dry` 模拟运行，不真实发送

## 权限说明

**macOS：** 需在「系统设置 → 隐私与安全性 → 辅助功能」中授予 Terminal 或 Python 权限。

**Windows：** 首次运行 uiautomation 时需以管理员身份运行，并允许 UI 自动化访问。

## 注意事项

- macOS / Windows 以外系统仅支持查看/配置功能
- 建议发送间隔不低于 3 秒，每分钟不超过 20 条
- 同一内容反复发送给多人容易被微信限制，建议适当变换文字内容
- 图片路径请使用绝对路径
