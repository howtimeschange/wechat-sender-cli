# 微信批量发送助手 CLI

跨平台命令行工具，支持 macOS 和 Windows。macOS 版可执行微信自动化发送，Windows 版可配置和查看任务状态。

## 系统要求

- **macOS**: Python 3.10+，支持微信自动化发送
- **Windows**: Python 3.10+，仅支持查看/配置（自动化发送需 macOS）

## 快速安装

### macOS

```bash
bash <(curl -fsSL https://raw.githubusercontent.com/你的用户名/wechat-sender-cli/main/install.sh)
```

或者手动：

```bash
git clone https://github.com/你的用户名/wechat-sender-cli.git
cd wechat-sender-cli
bash install.sh
```

### Windows

```cmd
git clone https://github.com/你的用户名/wechat-sender-cli.git
cd wechat-sender-cli
install.bat
```

## 依赖

- Python 3.10+
- openpyxl
- PyYAML
- rich

## 使用方法

```bash
# 首次配置（交互向导）
python3 app/cli.py setup

# 查看当前配置
python3 app/cli.py config

# 查看任务状态
python3 app/cli.py status

# 立即发送待处理任务（仅 macOS）
python3 app/cli.py send

# 守护进程模式（后台轮询，仅 macOS）
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
| excel_path | 表格文件路径 | - |
| poll_seconds | 守护进程轮询间隔（秒） | 15 |
| send_interval | 两条消息间隔（秒） | 5 |
| max_per_minute | 每分钟最大条数 | 8 |
| dry_run | 模拟运行（不真实发送） | false |

## 表格模板

| 列名 | 说明 | 示例 |
|------|------|------|
| * 应用 | 固定填 微信 | 微信 |
| * 联系人/群聊 | 对方微信名或群名 | 张三 |
| * 消息类型 | 文字/图片/文字+图片 | 文字 |
| * 文字内容 | 要发送的文本 | 你好！ |
| 图片路径 | 本地图片绝对路径 | /Users/xx/pic.png |
| 发送时间 | 格式 2025-01-01 14:30 | 2025-01-01 14:30 |
| 重复 | daily/weekly/workday/空 | daily |
| 状态 | 自动填充 | 发送成功 |

## 注意事项

- macOS 版需要授予「辅助功能」权限（终端或 Python）才能自动化操作微信
- 建议发送间隔不低于 3 秒，每分钟不超过 20 条
- 文字+图片需同时填写文字内容和图片路径
- 同一内容发给多人容易被微信限制，建议适当变换文字内容
