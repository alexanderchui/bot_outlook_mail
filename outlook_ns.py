import os
import re
import json
import time
import html
import traceback
import threading
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from urllib.parse import urlparse

from flask import Flask, request, jsonify
from O365 import Account, FileSystemTokenBackend


# ================= 核心配置区 =================
# 所有敏感信息均通过环境变量读取，部署时请自行设置

# 微软 OAuth 授权配置
CLIENT_ID = os.environ.get('OUTLOOK_CLIENT_ID', '')
CLIENT_SECRET = os.environ.get('OUTLOOK_CLIENT_SECRET', '')
credentials = (CLIENT_ID, CLIENT_SECRET)
SCOPES = ['User.Read', 'Mail.ReadWrite', 'Mail.Send']

# Web 服务配置
PORT = int(os.environ.get('OUTLOOK_PORT', '16666'))
AUTH_REDIRECT_URI = os.environ.get('OUTLOOK_REDIRECT_URI', f'http://localhost:{PORT}/callback')

# Token 存储目录（默认与本脚本同级的 outlook/ 文件夹）
TOKEN_DIR = os.environ.get('OUTLOOK_TOKEN_DIR', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outlook'))

# Telegram Bot 配置
TG_BOT_TOKEN = os.environ.get('TG_BOT_TOKEN', '')
TG_CHAT_ID = os.environ.get('TG_CHAT_ID', '')

# 推送引擎高级配置
TG_POLL_INTERVAL = 60                   # 默认轮询间隔（秒）
TG_PUSH_INTERVAL = 60                   # 推送间隔（秒）
TG_PUSH_FETCH_LIMIT = 300               # 每次扫描最大获取量
TG_PUSH_OVERLAP_SECONDS = 600           # 时间回看防漏推窗口（10分钟）
TG_PUSH_SEEN_ID_LIMIT = 2000            # 去重记忆池容量
TG_PUSH_FIRST_START_MINUTES = 0         # 首次启动补推时长（0为仅记录基线，不补推历史）
TG_PUSH_SEND_PAUSE = 0.35               # 连续推送防限流停顿时间（秒）

# ================= 状态文件与全局实例 =================

os.makedirs(TOKEN_DIR, mode=0o700, exist_ok=True)

# 用于在内存中传递隐藏的 MSAL Flow 字典
state_storage = {}
# 推送状态持久化文件
PUSH_STATE_PATH = os.path.join(TOKEN_DIR, 'tg_push_state.json')
# 屏蔽内容规则库
BLOCKED_CONTENTS_PATH = os.path.join(TOKEN_DIR, 'null.json')

app = Flask(__name__)


# ================= 核心工具函数 =================

def safe_account_name(account_name):
    """确保账号名合法，防止路径穿越风险"""
    account_name = (account_name or '').strip()
    if not re.fullmatch(r'[A-Za-z0-9_.@+-]{1,128}', account_name):
        raise ValueError("账号名只能包含字母、数字、下划线、点、@、+ 和横线，长度 1-128")
    return account_name

def _safe_html(text):
    if not text:
        return ''
    return html.escape(str(text), quote=False)

def _utcnow():
    return datetime.now(timezone.utc)

def _as_utc(dt):
    if not dt:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)

def _dt_to_iso(dt):
    dt = _as_utc(dt)
    return dt.isoformat() if dt else None

def _iso_to_dt(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace('Z', '+00:00')).astimezone(timezone.utc)
    except Exception:
        return None

def _fixed_callback_url(current_url):
    """修正 O365 回调时的 redirect_uri 校验问题"""
    try:
        target = urlparse(AUTH_REDIRECT_URI)
        current = urlparse(current_url)
        return current._replace(scheme=target.scheme, netloc=target.netloc).geturl()
    except Exception:
        return current_url


# ================= 速率限制器 =================

class RateLimiter:
    """简易滑动窗口限流器"""
    def __init__(self):
        self._buckets = defaultdict(list)
        self._lock = threading.Lock()

    def is_allowed(self, key, max_calls, period):
        now = time.time()
        with self._lock:
            bucket = self._buckets[key]
            self._buckets[key] = [t for t in bucket if now - t < period]
            if len(self._buckets[key]) >= max_calls:
                return False
            self._buckets[key].append(now)
            return True

rate_limiter = RateLimiter()

def check_rate(key, max_calls, period):
    if not rate_limiter.is_allowed(key, max_calls, period):
        return jsonify({'error': f'请求过于频繁，请 {period} 秒后再试'}), 429
    return None


# ================= O365 账户管理 =================

def get_authenticated_account(account_name):
    account_name = safe_account_name(account_name)
    token_filename = f'{account_name}.txt'
    token_backend = FileSystemTokenBackend(token_path=TOKEN_DIR, token_filename=token_filename)
    acc = Account(credentials, token_backend=token_backend)
    if not acc.is_authenticated:
        return None
    return acc

def get_all_account_names():
    files = os.listdir(TOKEN_DIR) if os.path.exists(TOKEN_DIR) else []
    accounts = []
    for f in files:
        if f.endswith('.txt'):
            try:
                accounts.append(safe_account_name(f[:-4]))
            except ValueError:
                pass
    return sorted(set(accounts))


# ================= Flask Web API =================

@app.route('/')
def index():
    ui_path = os.path.join(TOKEN_DIR, 'outlook_ui.html')
    if not os.path.exists(ui_path):
        return '<body style="background:#111;color:#eee;font-family:monospace;padding:40px;"><h2>Outlook Mail Center</h2><p>未找到 outlook_ui.html</p></body>', 404
    with open(ui_path, 'r', encoding='utf-8') as f:
        return f.read()

@app.route('/api/auth/<account_name>')
def auth(account_name):
    blocked = check_rate('auth_global', 3, 60)
    if blocked: return blocked
    try:
        account_name = safe_account_name(account_name)
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400
        
    print(f"[后端] 开始处理 [{account_name}] 的授权请求")
    try:
        token_filename = f'{account_name}.txt'
        token_backend = FileSystemTokenBackend(token_path=TOKEN_DIR, token_filename=token_filename)
        acc = Account(credentials, token_backend=token_backend)

        flow = acc.connection.msal_client.initiate_auth_code_flow(
            scopes=SCOPES, redirect_uri=AUTH_REDIRECT_URI, prompt='select_account'
        )
        url = flow.get('auth_uri')
        if not url: raise Exception("MSAL 未能生成授权 URL")

        state_str = flow.get('state', '')
        if not state_str: raise Exception("MSAL 未返回 state")

        state_storage[state_str] = {'account_name': account_name, 'flow': flow, 'created_at': time.time()}
        return jsonify({'status': 'success', 'auth_url': url})
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/callback')
def callback():
    fixed_url = _fixed_callback_url(request.url)
    state_str = request.args.get('state')
    stored_data = state_storage.get(state_str)

    if not stored_data:
        return "授权会话已过期，请重新在面板点击授权。", 400

    account_name = stored_data['account_name']
    hidden_flow_dict = stored_data['flow']
    token_backend = FileSystemTokenBackend(token_path=TOKEN_DIR, token_filename=f'{account_name}.txt')
    acc = Account(credentials, token_backend=token_backend)

    try:
        if acc.connection.request_token(fixed_url, flow=hidden_flow_dict):
            print(f"[后端] ✅ 账号 [{account_name}] 授权成功，Token 已落地")
            state_storage.pop(state_str, None)
            return """<script>if(window.opener){window.opener.postMessage('refresh', '*');window.close();}else{document.write('<body style="background:#000;color:#0f0;font-family:monospace;padding:50px;"><h2>[SYSTEM] AUTHENTICATION COMPLETE</h2><p>Token successfully written.</p></body>');}</script>"""
        return "授权失败：微软返回了空结果。", 500
    except Exception as e:
        print(f"[后端] ❌ 严重错误:\n{traceback.format_exc()}")
        return f"Auth Error: {str(e)}", 500

@app.route('/api/list')
def list_accounts():
    blocked = check_rate('list_global', 30, 60)
    if blocked: return blocked
    return jsonify({'code': 200, 'active_accounts': get_all_account_names()})

@app.route('/api/mail/<account_name>/<folder_type>')
def get_mail(account_name, folder_type):
    try: account_name = safe_account_name(account_name)
    except ValueError as e: return jsonify({'error': str(e)}), 400

    blocked = check_rate(f'mail_{account_name}', 10, 60)
    if blocked: return blocked

    acc = get_authenticated_account(account_name)
    if not acc: return jsonify({'error': '账号未认证，请重新授权'}), 401

    try:
        mailbox = acc.mailbox()
        folder = mailbox.inbox_folder() if folder_type == 'inbox' else mailbox.junk_folder() if folder_type == 'junk' else None
        if not folder: return jsonify({'error': '无效的文件夹类型'}), 400

        cutoff = _utcnow() - timedelta(days=30)
        messages = []

        for msg in folder.get_messages(limit=300, order_by='receivedDateTime desc', batch=50):
            msg_received = _as_utc(getattr(msg, 'received', None))
            if msg_received and msg_received < cutoff: break
            
            messages.append({
                'id': msg.object_id,
                'subject': msg.subject or '(无主题)',
                'sender_name': getattr(msg.sender, 'name', '') if msg.sender else '',
                'sender_addr': getattr(msg.sender, 'address', '') if msg.sender else '',
                'received': msg_received.isoformat() if msg_received else '',
                'preview': (msg.body_preview or '')[:120],
                'is_read': msg.is_read,
            })
        return jsonify({'messages': messages})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/mail/<account_name>/send', methods=['POST'])
def send_mail(account_name):
    try: account_name = safe_account_name(account_name)
    except ValueError as e: return jsonify({'error': str(e)}), 400

    blocked = check_rate(f'send_{account_name}', 5, 60)
    if blocked: return blocked

    acc = get_authenticated_account(account_name)
    if not acc: return jsonify({'error': '账号未认证，请重新授权'}), 401

    data = request.json
    if not data or not data.get('to') or not data.get('subject'):
        return jsonify({'error': '收件人和主题不能为空'}), 400

    try:
        m = acc.new_message()
        for addr in data['to'].split(','):
            addr = addr.strip()
            if addr: m.to.add(addr)
        m.subject = data['subject']
        m.body = data.get('body', '')
        m.send()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ================= Telegram Bot 中枢逻辑 =================

_user_states = {}
_push_state_lock = threading.Lock()
_bot_send_lock = threading.Lock()


def start_tg_bot():
    if not TG_BOT_TOKEN:
        print("[TG Bot] ⚠️ 未配置 TG_BOT_TOKEN，Bot 不启动")
        return

    try:
        import telebot
        from telebot.types import BotCommand, InlineKeyboardMarkup, InlineKeyboardButton
    except ImportError:
        print("[TG Bot] ⚠️ pyTelegramBotAPI 未安装，请运行: pip install pyTelegramBotAPI")
        return

    print("[TG Bot] 🌐 当前使用网络直连模式启动")
    bot = telebot.TeleBot(TG_BOT_TOKEN, parse_mode='HTML')

    def is_authorized_chat(chat_id):
        return not TG_CHAT_ID or str(chat_id) == str(TG_CHAT_ID)

    def is_authorized_message(message):
        return message and is_authorized_chat(message.chat.id)

    # ---- 菜单与设置 ----
    try:
        bot.set_my_commands([
            BotCommand("start", "🚀 所有指令"),
            BotCommand("l", "📋 账户列表"),
            BotCommand("s", "✉️ 发送邮件"),
            BotCommand("v", "📬 查看今天邮箱"),
            BotCommand("v3", "📬 查看最近 3 天"),
            BotCommand("v7", "📬 查看最近 7 天"),
            BotCommand("x", "🛡️ 屏蔽内容"),
            BotCommand("cancel", "❌ 取消操作"),
        ])
    except Exception as e:
        print(f"[TG Bot] ⚠️ 快捷菜单初始化失败: {e}")

    # ---- 黑名单控制 ----
    blocked_contents = []
    if os.path.exists(BLOCKED_CONTENTS_PATH):
        try:
            with open(BLOCKED_CONTENTS_PATH, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            if isinstance(loaded, list):
                blocked_contents.extend([str(x) for x in loaded if str(x)])
        except Exception as e:
            print(f"[TG Bot] ⚠️ 读取屏蔽内容失败: {e}")

    def save_blocked_contents():
        try:
            with open(BLOCKED_CONTENTS_PATH, 'w', encoding='utf-8') as f:
                json.dump(blocked_contents, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[TG Bot] 保存 null.json 失败: {e}")

    def _clean_mail_body(msg, max_len=800):
        try:
            body = getattr(msg, 'body', '') or ''
            if not body.strip():
                body = getattr(msg, 'body_preview', '') or ''
                
            body = re.sub(r'<head[^>]*>.*?</head>', '', body, flags=re.IGNORECASE | re.DOTALL)
            body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.IGNORECASE | re.DOTALL)
            body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.IGNORECASE | re.DOTALL)
            body = re.sub(r'</?(div|p|br|tr|td|li|h[1-6]|table|tbody|thead|section|article)[^>]*>', '\n', body, flags=re.IGNORECASE)
            body = re.sub(r'<[^>]+>', '', body)
            body = html.unescape(body)
            body = re.sub(r'[\u200b-\u200f\ufeff\u034f]', '', body)
            body = body.replace('\xa0', ' ').replace('\r', '')

            for b_text in blocked_contents:
                if b_text: body = body.replace(b_text, '')

            body = '\n'.join(line.strip() for line in body.split('\n') if line.strip())

            if not body.strip():
                body = getattr(msg, 'body_preview', '') or ''
                body = html.unescape(body)
                body = re.sub(r'[\u200b-\u200f\ufeff\u034f]', '', body)
                for b_text in blocked_contents:
                    if b_text: body = body.replace(b_text, '')
                body = '\n'.join(line.strip() for line in body.split('\n') if line.strip())

            if len(body) > max_len:
                body = body[:max_len] + '...(内容过长已截断)'
                
            body = _safe_html(body)
        except Exception:
            body = getattr(msg, 'body_preview', '') or ''
            for b_text in blocked_contents:
                if b_text: body = body.replace(b_text, '')
            body = _safe_html(body[:max_len])

        return body if body.strip() else '(无内容)'

    # ---- 交互指令 ----
    @bot.message_handler(commands=['start'])
    def cmd_start(message):
        if not is_authorized_message(message): return
        bot.send_message(message.chat.id, "📮 <b>Outlook 邮件中枢 Bot</b>\n━━━━━━━━━━━━━━━\n\n可用指令：\n\n/l  — 📋 账户列表\n/s  — ✉️ 发送邮件\n/v  — 📬 查看今天邮箱\n/v3 — 📬 查看最近 3 天\n/v7 — 📬 查看最近 7 天\n/x  — 🛡️ 屏蔽内容\n/cancel — ❌ 取消当前操作")

    @bot.message_handler(commands=['l'])
    def cmd_list(message):
        if not is_authorized_message(message): return
        accs = get_all_account_names()
        if not accs: return bot.send_message(message.chat.id, "📋 暂无已添加的 Outlook 账户")
        lines = ["📋 <b>已添加的 Outlook 账户：</b>", ""]
        for i, a in enumerate(accs, 1): lines.append(f"{i}. <code>{_safe_html(a)}</code> ✅")
        bot.send_message(message.chat.id, "\n".join(lines) + f"\n\n共 {len(accs)} 个账户")

    @bot.message_handler(commands=['cancel'])
    def cmd_cancel(message):
        if not is_authorized_message(message): return
        _user_states.pop(message.from_user.id, None)
        bot.send_message(message.chat.id, "❌ 已取消")

    @bot.message_handler(commands=['x'])
    def cmd_block(message):
        if not is_authorized_message(message): return
        _user_states[message.from_user.id] = {'step': 'input_block'}
        bot.send_message(message.chat.id, "🛡️ <b>屏蔽内容添加</b>\n\n请回复你需要屏蔽的文本内容（精准匹配）：")

    @bot.message_handler(func=lambda m: _user_states.get(m.from_user.id, {}).get('step') == 'input_block')
    def cmd_block_input(message):
        if not is_authorized_message(message): return
        b_text = (message.text or '').strip()
        _user_states.pop(message.from_user.id, None)
        if not b_text: return bot.send_message(message.chat.id, "❌ 内容为空，已取消")
        if b_text not in blocked_contents:
            blocked_contents.append(b_text)
            save_blocked_contents()
        bot.send_message(message.chat.id, f"✅ 已成功添加屏蔽模板。\n\n当前屏蔽列表共 {len(blocked_contents)} 条记录。")

    @bot.message_handler(commands=['s'])
    def cmd_send(message):
        if not is_authorized_message(message): return
        accs = get_all_account_names()
        if not accs: return bot.send_message(message.chat.id, "❌ 暂无可用账号，请先在网页端添加")
        markup = InlineKeyboardMarkup()
        for i, a in enumerate(accs): markup.add(InlineKeyboardButton(a, callback_data=f"sacc_{i}"))
        _user_states[message.from_user.id] = {'step': 'choose_acc', 'accs': accs}
        bot.send_message(message.chat.id, "📮 请选择发送账号：", reply_markup=markup)

    @bot.callback_query_handler(func=lambda c: c.data.startswith('sacc_'))
    def send_choose_acc(call):
        if not is_authorized_chat(call.message.chat.id): return
        state = _user_states.get(call.from_user.id)
        if not state or state.get('step') != 'choose_acc': return
        try: idx = int(call.data.split('_', 1)[1])
        except Exception: return
        accs = state.get('accs', [])
        if idx >= len(accs):
            bot.edit_message_text("❌ 无效选择", call.message.chat.id, call.message.message_id)
            _user_states.pop(call.from_user.id, None)
            return
        state.update({'send_acc': accs[idx], 'step': 'input_to'})
        bot.edit_message_text(f"✅ 已选择: <code>{_safe_html(accs[idx])}</code>\n\n📬 请输入收件人邮箱：", call.message.chat.id, call.message.message_id)

    @bot.message_handler(func=lambda m: _user_states.get(m.from_user.id, {}).get('step') == 'input_to')
    def send_input_to(message):
        if not is_authorized_message(message): return
        _user_states[message.from_user.id].update({'send_to': (message.text or '').strip(), 'step': 'input_subject'})
        bot.send_message(message.chat.id, "📝 请输入邮件主题：")

    @bot.message_handler(func=lambda m: _user_states.get(m.from_user.id, {}).get('step') == 'input_subject')
    def send_input_subject(message):
        if not is_authorized_message(message): return
        _user_states[message.from_user.id].update({'send_subject': (message.text or '').strip(), 'step': 'input_body'})
        bot.send_message(message.chat.id, "✏️ 请输入邮件内容：")

    @bot.message_handler(func=lambda m: _user_states.get(m.from_user.id, {}).get('step') == 'input_body')
    def send_input_body(message):
        if not is_authorized_message(message): return
        d = _user_states[message.from_user.id]
        d.update({'send_body': message.text or '', 'step': 'confirm'})
        bot.send_message(message.chat.id, f"📤 <b>确认发送？</b>\n\n👤 <b>发件人:</b> <code>{_safe_html(d['send_acc'])}</code>\n🎯 <b>收件人:</b> <code>{_safe_html(d['send_to'])}</code>\n📌 <b>主  题:</b> {_safe_html(d['send_subject'])}\n\n请回复 <b>Y</b> 确认发送，<b>N</b> 取消")

    @bot.message_handler(func=lambda m: _user_states.get(m.from_user.id, {}).get('step') == 'confirm')
    def send_confirm(message):
        if not is_authorized_message(message): return
        d = _user_states.pop(message.from_user.id, {})
        if (message.text or '').strip().upper() != 'Y': return bot.send_message(message.chat.id, "❌ 已取消发送")
        acc = get_authenticated_account(d.get('send_acc', ''))
        if not acc: return bot.send_message(message.chat.id, "❌ 账号未认证，请在网页端重新授权")
        try:
            m = acc.new_message()
            for addr in d['send_to'].split(','):
                if addr.strip(): m.to.add(addr.strip())
            m.subject = d['send_subject']
            m.body = d.get('send_body', '')
            m.send()
            bot.send_message(message.chat.id, "✅ 邮件发送成功！")
        except Exception as e:
            bot.send_message(message.chat.id, f"❌ 发送失败: {_safe_html(e)}")

    @bot.message_handler(commands=['v', 'v3', 'v7'])
    def cmd_view(message):
        if not is_authorized_message(message): return
        accs = get_all_account_names()
        if not accs: return bot.send_message(message.chat.id, "❌ 暂无可用账号")
        
        cmd = (message.text or '').split()[0].split('@')[0].strip()
        days = 3 if cmd == '/v3' else 7 if cmd == '/v7' else 1
        
        markup = InlineKeyboardMarkup()
        for i, a in enumerate(accs): markup.add(InlineKeyboardButton(a, callback_data=f"vacc_{i}"))
        markup.add(InlineKeyboardButton("🌐 ━ 查看全部邮箱 ━ 🌐", callback_data="vacc_all"))
        
        _user_states[message.from_user.id] = {'step': 'view_acc', 'accs': accs, 'days': days}
        day_text = "今天" if days == 1 else f"最近 {days} 天"
        bot.send_message(message.chat.id, f"📬 请选择要查看的账号（获取 {day_text} 邮件）：", reply_markup=markup)

    @bot.callback_query_handler(func=lambda c: c.data.startswith('vacc_'))
    def view_choose_acc(call):
        if not is_authorized_chat(call.message.chat.id): return
        state = _user_states.pop(call.from_user.id, {})
        accs, days = state.get('accs') or get_all_account_names(), int(state.get('days', 1))
        target_data = call.data.split('_', 1)[1]

        if target_data == "all":
            target_accounts = accs
            bot.edit_message_text(f"⏳ 正在扫描所有 {len(accs)} 个账号的邮件...", call.message.chat.id, call.message.message_id)
        else:
            try: idx = int(target_data)
            except Exception: return bot.edit_message_text("❌ 无效选择", call.message.chat.id, call.message.message_id)
            if idx >= len(accs): return bot.edit_message_text("❌ 无效选择", call.message.chat.id, call.message.message_id)
            target_accounts = [accs[idx]]
            bot.edit_message_text(f"⏳ 正在加载 {_safe_html(target_accounts[0])} 的邮件...", call.message.chat.id, call.message.message_id)

        try:
            cutoff = _utcnow() - timedelta(days=days)
            all_segments = []
            for account_name in target_accounts:
                acc = get_authenticated_account(account_name)
                if not acc:
                    if target_data != "all": all_segments.append(f"❌ {_safe_html(account_name)} 账号未认证")
                    continue
                try:
                    mailbox = acc.mailbox()
                    for fname, flabel in [('inbox', '📥 收件箱'), ('junk', '🗑️ 垃圾邮件')]:
                        folder = mailbox.inbox_folder() if fname == 'inbox' else mailbox.junk_folder()
                        blocks = []
                        for msg in folder.get_messages(limit=300, order_by='receivedDateTime desc', batch=50):
                            msg_received = _as_utc(getattr(msg, 'received', None))
                            if msg_received and msg_received < cutoff: break
                            sender = getattr(msg.sender, 'address', '') or getattr(msg.sender, 'name', '') or '未知' if msg.sender else '未知'
                            ts = msg_received.strftime('%m-%d %H:%M') if msg_received else ''
                            subj = _safe_html(msg.subject or '(无主题)')
                            attach_str = " 📎" if getattr(msg, 'has_attachments', False) else ""
                            body = _clean_mail_body(msg, max_len=800)
                            blocks.append(f"👤 <b>发  件:</b> <code>{_safe_html(sender)}</code> ｜ 🕒 <b>{ts}</b>\n📌 <b>标  题:</b> {subj}{attach_str}\n📝 <b>内  容:</b>\n{body}")

                        header = f"📬 <b>{_safe_html(account_name)}</b>\n{flabel} ({len(blocks)}封)\n{'━' * 20}"
                        if not blocks:
                            if target_data != "all": all_segments.append(f"{header}\n(空)")
                            continue
                        
                        cur = header
                        for blk in blocks:
                            entry = f"\n\n{blk}"
                            if len(cur) + len(entry) > 3800:
                                all_segments.append(cur)
                                cur = f"{flabel}（续）\n{'━' * 20}\n\n{blk}"
                            else:
                                cur += entry
                        if cur: all_segments.append(cur)
                except Exception as e:
                    if target_data != "all": all_segments.append(f"❌ 获取 {_safe_html(account_name)} 失败: {_safe_html(e)}")

            chat_id, msg_id = call.message.chat.id, call.message.message_id
            title = "所有账号" if target_data == "all" else target_accounts[0]
            if all_segments:
                bot.edit_message_text(f"✅ {_safe_html(title)} 邮件加载完成", chat_id, msg_id)
                for seg in all_segments: bot.send_message(chat_id, seg, disable_web_page_preview=True)
            else:
                bot.edit_message_text(f"📭 {_safe_html(title)} 暂无新邮件", chat_id, msg_id)
        except Exception as e:
            try: bot.edit_message_text(f"❌ 获取失败: {_safe_html(e)}", call.message.chat.id, call.message.message_id)
            except Exception: pass

    # ================= 新邮件实时推送：持久化游标版 =================

    def _load_push_state():
        with _push_state_lock:
            if not os.path.exists(PUSH_STATE_PATH): return {'version': 1, 'accounts': {}}
            try:
                with open(PUSH_STATE_PATH, 'r', encoding='utf-8') as f: data = json.load(f)
                if not isinstance(data, dict): raise ValueError("state file is not a dict")
                data.setdefault('version', 1)
                data.setdefault('accounts', {})
                return data
            except Exception as e:
                print(f"[TG Bot] ⚠️ 读取推送状态失败: {e}")
                return {'version': 1, 'accounts': {}}

    def _save_push_state(state):
        with _push_state_lock:
            try:
                tmp_path = PUSH_STATE_PATH + '.tmp'
                with open(tmp_path, 'w', encoding='utf-8') as f: json.dump(state, f, ensure_ascii=False, indent=2)
                os.replace(tmp_path, PUSH_STATE_PATH)
            except Exception as e:
                print(f"[TG Bot] ⚠️ 保存推送状态失败: {e}")

    def _get_folder_state(state, account_name, folder_key):
        return state.setdefault('accounts', {}).setdefault(account_name, {}).setdefault(folder_key, {'last_received': None, 'seen_ids': []})

    def _remember_seen_id(folder_state, msg_id):
        if not msg_id: return
        ids = folder_state.setdefault('seen_ids', [])
        if msg_id not in ids: ids.append(msg_id)
        if len(ids) > TG_PUSH_SEEN_ID_LIMIT: del ids[:-TG_PUSH_SEEN_ID_LIMIT]

    def _send_new_mail_push(account_name, folder_label, msg):
        sender_raw = getattr(msg.sender, 'address', '') or getattr(msg.sender, 'name', '') or '未知' if getattr(msg, 'sender', None) else '未知'
        to_list = [getattr(t, 'address', '') for t in getattr(msg, 'to', []) if getattr(t, 'address', '')]
        
        msg_time = _as_utc(getattr(msg, 'received', None))
        ts = msg_time.strftime('%m-%d %H:%M') if msg_time else ''
        subj = _safe_html(msg.subject or '(无主题)')
        attach_str = " 📎" if getattr(msg, 'has_attachments', False) else ""
        body = _clean_mail_body(msg, max_len=300)

        text = (f"🔔 <b>实时新邮件到达 ({folder_label})</b> ｜ 🕒 <b>{ts}</b>\n"
                f"📮 <b>账  号:</b> <code>{_safe_html(account_name)}</code>\n"
                f"👤 <b>发  件:</b> <code>{_safe_html(sender_raw)}</code>\n"
                f"🎯 <b>收  件:</b> <code>{_safe_html(', '.join(to_list) or '未知')}</code>\n"
                f"📌 <b>标  题:</b> {subj}{attach_str}\n"
                f"📝 <b>内  容:</b>\n{body}")

        with _bot_send_lock:
            bot.send_message(TG_CHAT_ID, text, parse_mode='HTML', disable_web_page_preview=True)

    def _seed_folder_baseline(state, account_name, folder_key, folder_label, messages):
        folder_state = _get_folder_state(state, account_name, folder_key)
        latest_dt, seed_ids = None, []
        for msg in messages:
            if mid := getattr(msg, 'object_id', None): seed_ids.append(mid)
            if (msg_dt := _as_utc(getattr(msg, 'received', None))) and (latest_dt is None or msg_dt > latest_dt):
                latest_dt = msg_dt

        folder_state['last_received'] = _dt_to_iso(latest_dt or _utcnow())
        folder_state['seen_ids'] = list(dict.fromkeys(seed_ids))[-TG_PUSH_SEEN_ID_LIMIT:]
        _save_push_state(state)
        print(f"[TG Bot] 📌 初始化推送基线: {account_name}/{folder_key}/{folder_label}，记录 {len(seed_ids)} 封")

    def _scan_and_push_folder(state, account_name, folder_key, folder_label, folder):
        folder_state = _get_folder_state(state, account_name, folder_key)
        try: messages = list(folder.get_messages(limit=TG_PUSH_FETCH_LIMIT, order_by='receivedDateTime desc', batch=50))
        except Exception as e: return print(f"[TG Bot] ⚠️ 拉取 {account_name}/{folder_key} 失败: {e}")

        last_dt = _iso_to_dt(folder_state.get('last_received'))
        if last_dt is None:
            if TG_PUSH_FIRST_START_MINUTES <= 0: return _seed_folder_baseline(state, account_name, folder_key, folder_label, messages)
            last_dt = _utcnow() - timedelta(minutes=TG_PUSH_FIRST_START_MINUTES)
            folder_state['last_received'], folder_state['seen_ids'] = _dt_to_iso(last_dt), []

        cutoff_dt = last_dt - timedelta(seconds=TG_PUSH_OVERLAP_SECONDS)
        seen_ids = set(folder_state.get('seen_ids', []))
        candidates = []

        for msg in messages:
            if not (mid := getattr(msg, 'object_id', None)): continue
            msg_dt = _as_utc(getattr(msg, 'received', None)) or _utcnow()
            if msg_dt < cutoff_dt: break
            if mid not in seen_ids: candidates.append((msg_dt, mid, msg))

        candidates.sort(key=lambda x: (x[0], x[1]))
        
        for msg_dt, mid, msg in candidates:
            try:
                _send_new_mail_push(account_name, folder_label, msg)
                _remember_seen_id(folder_state, mid)
                if msg_dt > (_iso_to_dt(folder_state.get('last_received')) or last_dt):
                    folder_state['last_received'] = _dt_to_iso(msg_dt)
                _save_push_state(state)
                if TG_PUSH_SEND_PAUSE > 0: time.sleep(TG_PUSH_SEND_PAUSE)
            except Exception as e:
                print(f"[TG Bot] ⚠️ TG 推送失败: {account_name}/{folder_key} | {e}")
                _save_push_state(state)
                return

    def mail_poll_loop():
        if not TG_CHAT_ID: return print("[TG Bot] ⚠️ 未配置 TG_CHAT_ID，推送不启动")
        time.sleep(5)
        while True:
            cycle_started = time.time()
            try:
                state = _load_push_state()
                for acc_name in get_all_account_names():
                    try:
                        if acc := get_authenticated_account(acc_name):
                            mailbox = acc.mailbox()
                            for f_key, f_obj, f_label in [('inbox', mailbox.inbox_folder(), "📥收件箱"), ('junk', mailbox.junk_folder(), "🗑️垃圾箱")]:
                                _scan_and_push_folder(state, acc_name, f_key, f_label, f_obj)
                    except Exception as e_acc: print(f"[TG Bot] ⚠️ 账号 {acc_name} 轮询报错: {e_acc}")
                _save_push_state(state)
            except Exception as e_main: print(f"[TG Bot] ⚠️ 外层大循环报错: {e_main}")
            time.sleep(max(5, TG_PUSH_INTERVAL - (time.time() - cycle_started)))

    # ---- 守护线程运行逻辑 ----
    def run_bot():
        print("[TG Bot] ✅ Bot 已上线")
        while True:
            try: bot.infinity_polling(skip_pending=True, timeout=60, long_polling_timeout=60)
            except Exception as e:
                err = str(e)
                if any(x in err for x in ["Read timed out", "SSLZeroReturnError", "Connection aborted", "ConnectionError"]):
                    print("[TG Bot] ⏳ 网络波动，10 秒后自动重连...")
                elif "409" in err or "Conflict" in err:
                    print("[TG Bot] ⚠️ 检测到 409 Conflict，可能有另一个 Bot 实例正在运行，10 秒后重试...")
                else: print(f"[TG Bot] ❌ 轮询异常: {e}")
                time.sleep(10)

    def send_startup_notice():
        if not TG_CHAT_ID: return 
        try:
            with _bot_send_lock:
                bot.send_message(TG_CHAT_ID, "🚀 <b>系统启动成功</b>\n🌐 <b>当前使用网络直连模式</b>", parse_mode="HTML")
        except Exception as e: print(f"[TG Bot] 启动推送失败: {e}")

    threading.Thread(target=run_bot, daemon=True).start()
    threading.Thread(target=mail_poll_loop, daemon=True).start()
    threading.Thread(target=send_startup_notice, daemon=True).start()

    print("[TG Bot] 后台双核线程（监听/收信）已全部启动")


# ================= 启动 =================

if __name__ == '__main__':
    # 启动前检查必要配置
    if not CLIENT_ID or not CLIENT_SECRET:
        print("⚠️  请设置环境变量: OUTLOOK_CLIENT_ID 和 OUTLOOK_CLIENT_SECRET")
        print("   export OUTLOOK_CLIENT_ID='your-client-id'")
        print("   export OUTLOOK_CLIENT_SECRET='your-client-secret'")
        print("   (可选) export TG_BOT_TOKEN='your-bot-token'")
        print("   (可选) export TG_CHAT_ID='your-chat-id'")
        print("   (可选) export OUTLOOK_PORT='16666'")
        print("   (可选) export OUTLOOK_REDIRECT_URI='http://localhost:16666/callback'")
        print()
    
    start_tg_bot()
    app.run(host='0.0.0.0', port=PORT, debug=False, use_reloader=False)
