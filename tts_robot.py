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
    "volume": 1.0
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

def add_play(nickname, text, read_name):
    if not text:
        return
    content = f"{nickname}说：{text}" if read_name else text
    speak_queue.append({"content": content})
    if not is_speaking:
        play_next()

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
        self.read_name_var = tk.BooleanVar(value=True)

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
        ttk.Checkbutton(top, text="播报用户名", variable=self.read_name_var).grid(row=0, column=1, padx=5)

        # 音色
        ttk.Label(top, text="音色：").grid(row=0, column=2, padx=2)
        voice_var = tk.StringVar(value=voices[config["voice"]] if voices else "微软语音")
        voice_box = ttk.Combobox(top, textvariable=voice_var, values=voices, state="readonly", width=22)
        voice_box.grid(row=0, column=3, padx=5)
        voice_box.bind("<<ComboboxSelected>>", self.save_config)

        # 语速
        ttk.Label(top, text="语速：").grid(row=0, column=4, padx=2)
        rate_var = tk.IntVar(value=config["rate"])
        rate_scale = ttk.Scale(top, from_=-10, to=20, variable=rate_var)
        rate_scale.grid(row=0, column=5, padx=5)
        rate_scale.bind("<ButtonRelease-1>", self.save_config)

        # 音量
        ttk.Label(top, text="音量：").grid(row=0, column=6, padx=2)
        vol_var = tk.DoubleVar(value=config["volume"])
        vol_scale = ttk.Scale(top, from_=0.0, to=1.0, variable=vol_var)
        vol_scale.grid(row=0, column=7, padx=5)
        vol_scale.bind("<ButtonRelease-1>", self.save_config)

    # ==========================
    # 保存配置（滑块松开后保存）
    # ==========================
    def save_config(self, event=None):
        try:
            config["voice"] = voices.index(voice_var.get())
            config["rate"] = rate_var.get()
            config["volume"] = vol_var.get()

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
            ttk.Button(row, text="播放", command=lambda n=nickname, t=text: add_play(n, t, self.read_name_var.get())).pack(side="right", padx=4)

            add_play(nickname, text, self.read_name_var.get())
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