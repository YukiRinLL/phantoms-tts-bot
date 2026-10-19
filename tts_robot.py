import tkinter as tk
from tkinter import ttk, Frame, Canvas, Scrollbar, scrolledtext
import threading
import time
import re
import requests
import json
import subprocess
import os

# ==========================
# 配置文件
# ==========================
CONFIG_FILE = "tts_config.json"
config = {
    "voice": 0,
    "rate": 10,
    "volume": 1.0,
    "name_mode": "full"
}

try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
except:
    pass

# ==========================
# 加载系统语音
# ==========================
voices = []
def load_voices():
    global voices
    try:
        res = subprocess.check_output(
            ["powershell", "-Command", "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; $s.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name }"],
            creationflags=0x08000000, encoding="utf-8"
        )
        voices = [v.strip() for v in res.strip().splitlines() if v.strip()]
    except:
        voices = ["Microsoft Huihui Desktop"]

load_voices()

# ==========================
# 全局播报控制
# ==========================
speak_queue = []
is_speaking = False
message_ids = set()

# ==========================
# ✅ 播报函数（实时读取最新音量、语速）
# ==========================
def speak_sync(text):
    try:
        text = text.replace("'", "").replace('"', "")[:300]

        # 实时读取滑块当前值 → 立刻生效
        rate = config["rate"] - 10
        vol = int(config["volume"] * 100)
        voice_name = voices[config["voice"]] if (voices and config["voice"] < len(voices)) else "Microsoft Huihui Desktop"

        cmd = (
            f"Add-Type -AssemblyName System.Speech; "
            f"$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            f"$s.Rate = {rate}; "
            f"$s.Volume = {vol}; "
            f"$s.SelectVoice('{voice_name}'); "
            f"$s.Speak('{text}');"
        )
        subprocess.run(
            ["powershell", "-Command", cmd],
            creationflags=0x08000000,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception as e:
        print("播报错误：", e)

# ==========================
# 顺序播报队列
# ==========================
def play_next():
    global is_speaking
    if not speak_queue:
        is_speaking = False
        return

    is_speaking = True
    item = speak_queue.pop(0)
    content = item["content"]
    log(f"正在播报：{content}")

    def _run():
        global is_speaking
        speak_sync(content)
        time.sleep(0.2)
        is_speaking = False
        play_next()

    threading.Thread(target=_run, daemon=True).start()

def smart_name(nickname):
    """Keep the likely name while removing common QQ nickname decorations."""
    name = re.sub(r"[\s\u200b\u200c\u200d\ufeff]+", "", str(nickname or "用户"))

    # QQ nicknames commonly use these paired marks for a clan, role, status,
    # platform tag, or fan-group label before/after the actual name.
    bracket_pairs = (
        ("【", "】"), ("［", "］"), ("[", "]"), ("（", "）"), ("(", ")"),
        ("{", "}"), ("｛", "｝"), ("<", ">"), ("《", "》"), ("〈", "〉"),
        ("「", "」"), ("『", "』"), ("〔", "〕"), ("〖", "〗"),
        ("〘", "〙"), ("〚", "〛"), ("﹙", "﹚"),
    )
    changed = True
    while changed and name:
        changed = False
        for left, right in bracket_pairs:
            prefix = re.match(re.escape(left) + r".*?" + re.escape(right), name)
            if prefix:
                name = name[prefix.end():]
                changed = True

            suffix = re.search(re.escape(left) + r".*?" + re.escape(right) + r"$", name)
            if suffix:
                name = name[:suffix.start()]
                changed = True

    # Split only on separators normally used between a name and a tag. Keep
    # underscores out because they are common in genuine English usernames.
    separators = r"[-·•‧・･/／\\|｜¦丨—–~～=＝:：,，、]+"
    parts = [part for part in re.split(separators, name) if part]

    # Strip unpaired decoration frequently used around QQ display names.
    edge_marks = r"★☆✦✧✩✪✫✬✭✮✯✰♡♥❤ღ❀✿❁❃☀☾☽♔♕♛♚⚡☯※◆◇■□●○◎▲△▼▽→←➤➜➳➵➸➹➷➶" \
        r"`~!！@#$%^&*＊+=|｜:：;；,，。．…·•"
    name = re.sub(r"^[" + edge_marks + r"]+|[" + edge_marks + r"]+$", "", name)
    parts = [re.sub(r"^[" + edge_marks + r"]+|[" + edge_marks + r"]+$", "", part) for part in parts]
    parts = [part for part in parts if part]

    common_tags = re.compile(
        r"(?:官方|认证|認證|客服|管理员|管理員|群主|房主|站长|站長|频道|頻道|工作室|团队|團隊|"
        r"备用|備用|小号|小號|主号|主號|马甲|馬甲|本尊|本人|粉丝|粉絲|后援|後援|应援|應援|"
        r"搬运|搬運|同人|游戏|遊戲|电竞|電競|直播|主播|店铺|店鋪|店主|合作|商务|商務|"
        r"bot|robot|admin|mod|vip|official|live|fan|fans|team|studio|club|tv|频道主)",
        re.IGNORECASE,
    )

    def is_decoration(part):
        if common_tags.search(part):
            return True
        if re.fullmatch(r"[A-Za-z0-9+_.]+", part) and any(re.search(r"[\u3400-\u9fff]", p) for p in parts):
            return True
        # Short Chinese edge labels such as “小号”“工作号” are frequent tags.
        if re.search(r"(?:号|號|版|站|组|組|团|團|部|区|區|服|台|社)$", part) and len(part) <= 4:
            return True
        return False

    while len(parts) > 1:
        first_tag = is_decoration(parts[0])
        last_tag = is_decoration(parts[-1])
        if first_tag and not last_tag:
            parts.pop(0)
        elif last_tag and not first_tag:
            parts.pop()
        elif first_tag and last_tag:
            parts.pop(0)
            if len(parts) > 1:
                parts.pop()
        else:
            # For two Chinese fragments, a suffix is more often a role/tag;
            # retain the first fragment (e.g. “小明-后援会”).
            if len(parts) == 2 and all(re.search(r"[\u3400-\u9fff]", p) for p in parts):
                parts.pop()
            else:
                break

    name = parts[0] if parts else name
    return name or "用户"


def add_play(nickname, text, name_mode):
    if not text:
        return
    content = build_speech(nickname, text, name_mode)
    speak_queue.append({"content": content})
    if not is_speaking:
        play_next()


def get_name_mode(display_name):
    return {
        "不播报用户名": "none",
        "完整用户名": "full",
        "智能裁剪": "smart",
        "简称（前两个字）": "short",
    }.get(display_name, "full")


def build_speech(nickname, text, name_mode):
    if name_mode == "none":
        return text
    if name_mode == "full":
        return f"{nickname}说：{text}"
    if name_mode == "smart":
        return f"{smart_name(nickname)}，{text}"
    return f"{str(nickname or '用户')[:2]}，{text}"

# ==========================
# 跳过播报（正常可用）
# ==========================
def skip_play():
    global is_speaking
    is_speaking = False
    speak_queue.clear()
    log("===== 已跳过播报并清空队列 =====")

# ==========================
# 详细日志
# ==========================
def log(msg):
    try:
        time_str = time.strftime("%Y-%m-%d %H:%M:%S")
        log_box.insert(tk.END, f"[{time_str}] {msg}\n")
        log_box.see(tk.END)
    except:
        pass

# ==========================
# 主界面
# ==========================
class TTSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phantoms群消息TTS自动播报 v1.0.0.beta")
        self.root.geometry("960x780")

        # 图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "icon.jpg")
            if os.path.exists(icon_path):
                self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
        except:
            pass

        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.API_URL = "https://phantoms-backend.onrender.com/onebot/latest?limit=30"
        name_mode = config.get("name_mode", "full")
        if name_mode not in ("none", "full", "smart", "short"):
            name_mode = "full"
        name_mode_labels = {
            "none": "不播报用户名",
            "full": "完整用户名",
            "smart": "智能裁剪",
            "short": "简称（前两个字）",
        }
        self.name_mode_var = tk.StringVar(value=name_mode_labels[name_mode])

        self.create_top()
        self.create_list()
        self.create_log()
        log("程序启动成功")
        self.start_fetch()

    # ==========================
    # 顶部控制面板
    # ==========================
    def create_top(self):
        global rate_var, vol_var, voice_var
        top = Frame(self.root, padx=10, pady=6)
        top.grid(row=0, column=0, sticky="ew")

        ttk.Button(top, text="跳过播报", command=skip_play).grid(row=0, column=0, padx=5)
        ttk.Label(top, text="用户名：").grid(row=0, column=1, padx=2)
        name_mode_box = ttk.Combobox(
            top,
            textvariable=self.name_mode_var,
            values=("不播报用户名", "完整用户名", "智能裁剪", "简称（前两个字）"),
            state="readonly",
            width=12,
        )
        name_mode_box.grid(row=0, column=2, padx=5)
        name_mode_box.bind("<<ComboboxSelected>>", self.save_config)

        # 音色
        ttk.Label(top, text="音色：").grid(row=0, column=3, padx=2)
        voice_var = tk.StringVar(value=voices[config["voice"]] if voices else "微软语音")
        voice_box = ttk.Combobox(top, textvariable=voice_var, values=voices, state="readonly", width=22)
        voice_box.grid(row=0, column=4, padx=5)
        voice_box.bind("<<ComboboxSelected>>", self.save_config)

        # 语速
        ttk.Label(top, text="语速：").grid(row=0, column=5, padx=2)
        rate_var = tk.IntVar(value=config["rate"])
        rate_scale = ttk.Scale(top, from_=-10, to=20, variable=rate_var)
        rate_scale.grid(row=0, column=6, padx=5)
        rate_scale.bind("<ButtonRelease-1>", self.save_config)

        # 音量
        ttk.Label(top, text="音量：").grid(row=0, column=7, padx=2)
        vol_var = tk.DoubleVar(value=config["volume"])
        vol_scale = ttk.Scale(top, from_=0.0, to=1.0, variable=vol_var)
        vol_scale.grid(row=0, column=8, padx=5)
        vol_scale.bind("<ButtonRelease-1>", self.save_config)

    # ==========================
    # 保存配置（滑块松开后保存）
    # ==========================
    def save_config(self, event=None):
        try:
            if voice_var.get() in voices:
                config["voice"] = voices.index(voice_var.get())
            config["rate"] = rate_var.get()
            config["volume"] = vol_var.get()
            config["name_mode"] = {
                "不播报用户名": "none",
                "完整用户名": "full",
                "智能裁剪": "smart",
                "简称（前两个字）": "short",
            }.get(self.name_mode_var.get(), self.name_mode_var.get())

            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)

            log(f"配置已更新 → 语速:{config['rate']} 音量:{config['volume']:.2f}")
        except:
            pass

    # ==========================
    # 消息列表
    # ==========================
    def create_list(self):
        self.win = Frame(self.root)
        self.win.grid(row=1, column=0, sticky="nsew", padx=8, pady=5)
        self.win.rowconfigure(0, weight=1)
        self.win.columnconfigure(0, weight=1)

        self.canvas = Canvas(self.win)
        self.scroll = Scrollbar(self.win, orient="vertical", command=self.canvas.yview)
        self.box = Frame(self.canvas)

        self.box.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.box, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scroll.grid(row=0, column=1, sticky="ns")

    def create_log(self):
        global log_box
        f = Frame(self.root)
        f.grid(row=2, column=0, sticky="ew", padx=8, pady=4)
        f.columnconfigure(0, weight=1)
        log_box = scrolledtext.ScrolledText(f, height=6)
        log_box.grid(row=0, column=0, sticky="ew")

    # ==========================
    # 消息解析
    # ==========================
    def parse_msg(self, msg):
        try:
            s = msg.get("message", "")
            if not s.startswith("{type=text"):
                return None
            match = re.search(r'text=([^}]+)}', s)
            return match.group(1).strip() if match else None
        except:
            return None

    # ==========================
    # 增量拉取
    # ==========================
    def fetch(self):
        log("开始拉取消息...")
        try:
            data = requests.get(self.API_URL, timeout=10).json()
            items = list(reversed(data))
            log(f"拉取完成，共 {len(items)} 条")
        except:
            log("拉取消息失败")
            return

        new_count = 0
        for item in items:
            mid = str(item.get("id", ""))
            if mid in message_ids:
                continue

            text = self.parse_msg(item)
            if not text:
                continue

            message_ids.add(mid)
            nickname = item.get("nickname", "用户")
            log(f"新消息：{nickname}：{text[:40]}...")

            row = Frame(self.box, padx=6, pady=5)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"{nickname}：{text}", wraplength=820, anchor="w").pack(side="left", fill="x", expand=True)
            ttk.Button(row, text="播放", command=lambda n=nickname, t=text: add_play(n, t, get_name_mode(self.name_mode_var.get()))).pack(side="right", padx=4)

            add_play(nickname, text, get_name_mode(self.name_mode_var.get()))
            new_count += 1

        if new_count > 0:
            log(f"本次新增 {new_count} 条消息")
        else:
            log("暂无新消息")

    def start_fetch(self):
        def loop():
            while True:
                try:
                    self.fetch()
                except:
                    pass
                time.sleep(5)
        threading.Thread(target=loop, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    TTSApp(root)
    root.mainloop()
