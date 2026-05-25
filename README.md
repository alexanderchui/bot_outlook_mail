依赖安装 
Python 依赖（3个）

pip3 install flask O365 pyTelegramBotAPI
包名	版本要求	用途
flask	≥2.0	Web 服务框架，提供 API 和 UI 面板
O365	≥2.0	微软 Graph API SDK，邮件收发和 OAuth
pyTelegramBotAPI	≥4.0	Telegram Bot 交互（可选，不装则 Bot 不启动）

目录结构
your-project/
├── outlook_ns.py          # 主程序
└── outlook/               # 数据目录（必须与 py 同级）
    ├── outlook_ui.html    # Web 前端面板
    └── null.json          # 屏蔽规则（初始为空 []）

配置环境变量  (必须填写)    
# 必填 — 微软 OAuth 凭证
export OUTLOOK_CLIENT_ID='your-microsoft-app-client-id'
export OUTLOOK_CLIENT_SECRET='your-microsoft-app-client-secret'

# 可选 — Telegram Bot（不填则 Bot 功能不启动）
export TG_BOT_TOKEN='your-telegram-bot-token'
export TG_CHAT_ID='your-telegram-chat-id'

# 可选 — 服务配置   端口必须跟AZ上的一致，可自定义
export OUTLOOK_PORT='16666'
export OUTLOOK_REDIRECT_URI='http://localhost:16666/callback'
export OUTLOOK_TOKEN_DIR='./outlook'

启动服务
python3 outlook_ns.py
启动后：

🌐 Web 面板: http://localhost:16666
🤖 Telegram Bot: 自动连接（如已配置）
📬 新邮件推送: 每 60 秒轮询自动推送
4. 添加 Outlook 账号
打开 http://localhost:16666
点击 "+ 添加账号"
输入 Outlook 邮箱地址
跳转微软登录页完成授权
授权成功后自动刷新账号列表

功能概览
功能	Web 面板	Telegram Bot
账号管理	✅ 添加/查看	✅ /l 列表
收件查看	✅ 收件箱/垃圾箱	✅ /v /v3 /v7
发送邮件	✅ 弹窗发送	✅ /s 交互式
新邮件推送	—	✅ 实时推送
内容屏蔽	—	✅ /x 添加规则
