# Outlook Notify System

基于 Microsoft Graph API 的 Outlook 邮件通知系统，支持 Web 面板管理、多账号授权、Telegram Bot 实时推送、邮件发送、垃圾箱查看、关键词屏蔽等功能。

---

# 功能特性

- ✅ Outlook 多账号 OAuth 授权登录
- ✅ 收件箱 / 垃圾箱邮件查看
- ✅ Web UI 邮件管理面板
- ✅ Telegram Bot 实时交互
- ✅ 新邮件自动推送提醒
- ✅ 支持 Telegram 快速发送邮件
- ✅ 支持关键词屏蔽规则
- ✅ 自动轮询 Outlook 新邮件
- ✅ 本地 Token 持久化保存

---

# 环境依赖

## Python 依赖

安装以下依赖：

```bash
pip3 install flask O365 pyTelegramBotAPI
```

## 依赖说明

| 包名 | 版本要求 | 用途 |
|---|---|---|
| flask | ≥ 2.0 | Web 服务框架，提供 API 与 UI 面板 |
| O365 | ≥ 2.0 | Microsoft Graph API SDK，用于 Outlook OAuth 与邮件操作 |
| pyTelegramBotAPI | ≥ 4.0 | Telegram Bot 交互（可选） |

> 不安装 `pyTelegramBotAPI` 时，Telegram Bot 功能将自动禁用。

---

# 项目结构

```text
your-project/
├── outlook_ns.py          # 主程序
└── outlook/               # 数据目录（必须与 py 文件同级）
    ├── outlook_ui.html    # Web 前端页面
    └── null.json          # 屏蔽规则（初始内容为 []）
```

---

# 环境变量配置

## 必填配置（Microsoft OAuth）

```bash
export OUTLOOK_CLIENT_ID='your-microsoft-app-client-id'
export OUTLOOK_CLIENT_SECRET='your-microsoft-app-client-secret'
```

---

## 可选配置（Telegram Bot）

```bash
export TG_BOT_TOKEN='your-telegram-bot-token'
export TG_CHAT_ID='your-telegram-chat-id'
```

> 不配置时，Telegram Bot 功能不会启动。

---

## 服务配置（可选）

```bash
export OUTLOOK_PORT='16666'
export OUTLOOK_REDIRECT_URI='http://localhost:16666/callback'
export OUTLOOK_TOKEN_DIR='./outlook'
```

### 参数说明

| 变量名 | 默认值 | 说明 |
|---|---|---|
| OUTLOOK_PORT | 16666 | Web 服务端口 |
| OUTLOOK_REDIRECT_URI | http://localhost:16666/callback | OAuth 回调地址 |
| OUTLOOK_TOKEN_DIR | ./outlook | Token 存储目录 |

> `OUTLOOK_PORT` 必须与 Azure App Registration 中配置的 Redirect URI 端口一致。

---

# 启动服务

运行：

```bash
python3 outlook_ns.py
```

启动成功后：

```text
🌐 Web 面板:
http://localhost:16666

🤖 Telegram Bot:
自动连接（如已配置）

📬 新邮件推送:
每 60 秒自动轮询
```

---

# 添加 Outlook 账号

## Web 授权流程

1. 打开：

```text
http://localhost:16666
```

2. 点击：

```text
+ 添加账号
```

3. 输入 Outlook 邮箱地址

4. 跳转微软登录页完成授权

5. 授权成功后自动返回并刷新账号列表

---

# Telegram Bot 指令

| 指令 | 功能 |
|---|---|
| `/l` | 查看已登录账号 |
| `/v` | 查看最新邮件 |
| `/v3` | 查看最近 3 封邮件 |
| `/v7` | 查看最近 7 封邮件 |
| `/s` | 交互式发送邮件 |
| `/x` | 添加屏蔽关键词规则 |

---

# 功能概览

| 功能 | Web 面板 | Telegram Bot |
|---|---|---|
| 账号管理 | ✅ | ✅ `/l` |
| 收件查看 | ✅ 收件箱 / 垃圾箱 | ✅ `/v` `/v3` `/v7` |
| 邮件发送 | ✅ Web 弹窗发送 | ✅ `/s` |
| 新邮件推送 | ❌ | ✅ 实时推送 |
| 内容屏蔽 | ❌ | ✅ `/x` |

---

# 屏蔽规则

屏蔽规则保存在：

```text
outlook/null.json
```

初始内容：

```json
[]
```

示例：

```json
[
  "spam",
  "广告",
  "unsubscribe"
]
```

当邮件标题或内容命中关键词时，将不会推送通知。

---

# Microsoft Azure 配置说明

需要在 Azure Portal 创建应用并配置：

## Redirect URI

```text
http://localhost:16666/callback
```

## API 权限（Microsoft Graph）

建议添加：

- Mail.Read
- Mail.Send
- offline_access
- User.Read

---

# 适用场景

- Outlook 邮件实时通知
- Telegram 邮件提醒
- 多账号邮件统一管理
- 服务器邮件监控
- 自动化邮件通知系统

---

# License

MIT License
