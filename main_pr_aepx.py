import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path, PureWindowsPath, PurePosixPath
import gzip
import os
import xml.etree.ElementTree as ET
from datetime import datetime
import requests
import json
import traceback

# ----------------------------
# 設定 / 常數
# ----------------------------
CONFIG_PATH = Path.home() / ".pr_compare_tool_config.json"  # 用來記住使用者偏好
LAMBDA_URL = "https://btgm2bmux2k4dtmgvq3bv4pdmy0muepe.lambda-url.ap-southeast-2.on.aws/"

# ----------------------------
# GUI 初始化
# ----------------------------
root = tk.Tk()
if sys.platform == "darwin":
    # macOS：避免安全警告
    root.createcommand('tk::mac::OpenDocument', lambda *args: None)
    root.createcommand('tk::mac::Quit', lambda *args: root.quit())
    root.config(menu=tk.Menu(root))
root.title("PR 專案素材比對工具")

# ----------------------------
# 偏好載入 / 儲存
# ----------------------------

def load_config() -> dict:
    """讀取本機設定檔，若不存在則回傳空 dict"""
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config():
    """根據勾選情況將資料寫入設定檔"""
    cfg = {}
    if remember_uid_var.get() and uid_entry.get().strip():
        cfg["line_uid"] = uid_entry.get().strip()
    if remember_output_var.get() and output_entry.get().strip():
        cfg["output_dir"] = output_entry.get().strip()
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print("⚠️ 儲存設定失敗：", e)

# ----------------------------
# LINE 相關
# ----------------------------

def send_to_lambda(text: str, uid: str):
    """將訊息傳送給 AWS Lambda，由 Lambda 依照方案 A 發送 LINE push"""
    payload = {
        "text": text,
        "to_user_id": uid  # 方案 A：把 target uid 帶進 Lambda
    }
    try:
        response = requests.post(LAMBDA_URL, headers={'Content-Type': 'application/json'}, json=payload, timeout=15)
        print("✅ Lambda 回應：", response.text)
    except Exception as e:
        print("❌ Lambda 呼叫失敗：", e)

# ----------------------------
# 日誌設定（解析失敗時寫入 .log）
# ----------------------------
LOG_DIR = Path.home() / "pr_compare_logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

def write_error_log(project_label: str, exc: Exception) -> Path:
    """把完整 Traceback 寫入 log，回傳檔案路徑"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"{project_label}_{ts}.log"
    log_path.write_text(traceback.format_exc(), encoding="utf-8")
    return log_path

# ----------------------------
# .prproj 解析 + 比對工具
# ----------------------------

def _clean_to_filename(raw: str) -> str:
    s = raw.strip().strip('"').strip("'")
    if s.startswith('\\\\?\\'):
        s = s[4:]
    name = PureWindowsPath(s).name
    if not name or name == s:
        name = PurePosixPath(s).name
    return name

def parse_project_filenames(project_path: str | os.PathLike) -> set[str]:
    """
    根據副檔名，自動解析 .prproj（Premiere）或 .aepx（After Effects） 的素材檔案名稱
    """
    project_path = Path(project_path)
    suffix = project_path.suffix.lower()

    # .prproj 為 gzip 壓縮 XML
    if suffix == ".prproj" or project_path.suffixes[-2:] == ['.prproj', '.gz']:
        with project_path.open('rb') as f:
            magic = f.read(2)
            f.seek(0)
            content = (
                gzip.open(f, 'rt', encoding='utf-8').read()
                if magic == b'\x1f\x8b' else
                f.read().decode('utf-8-sig', errors='ignore')
            )
        root_xml = ET.fromstring(content)

    # .aepx 為純 XML（After Effects）
    elif suffix == ".aepx":
        root_xml = ET.parse(project_path).getroot()

    else:
        raise ValueError(f"不支援的專案格式：{suffix}")
    
    # 解析 XML：優先使用 lxml recover，失敗才退回內建 ET
    try:
        import lxml.etree as LET
        parser = LET.XMLParser(recover=True)
        root_xml = LET.fromstring(content.encode("utf-8"), parser=parser)
        iter_elems = root_xml.iter()
    except Exception:
        # 若 lxml 不存在或仍失敗，改用內建 ET（可能丟 ParseError）
        root_xml = ET.fromstring(content)
        iter_elems = root_xml.iter()
    
    candidate_tags = {"Path", "FilePath", "ActualMediaFilePath", "AbsolutePath"}
    filenames: set[str] = set()
    for elem in root_xml.iter():
        if elem.text and any(t.lower() in elem.tag.lower() for t in candidate_tags):
            name = _clean_to_filename(elem.text)
            if '.' in name:
                filenames.add(name)
    return filenames

def scan_folder_filenames(folder_path):
    folder_path = Path(folder_path)
    filenames = {f.name for f in folder_path.rglob('*') if f.is_file()}
    return filenames

IGNORE_EXTENSIONS = {".pek", ".epr", ".cfa"}          # 代理檔、預覽檔
IGNORE_SUBSTRINGS = {"_proxy", "_subclip"}            # 低畫質 Proxy / Subclip
# 忽略字串（不區分大小寫）
def is_ignored(filename: str) -> bool:
    """
    判斷該檔名是否屬於可忽略類型
    - 以副檔名為主；補充部分常見字串（不區分大小寫）
    """
    lower = filename.lower()
    if any(lower.endswith(ext) for ext in IGNORE_EXTENSIONS):
        return True
    if any(sub in lower for sub in IGNORE_SUBSTRINGS):
        return True
    return False
# 比對專案檔案與資料夾中的檔案名稱
# 添加檔案白名單.pek/.epr/proxy 
def compare_filenames(project_file, folder_path):
    project_files_raw = parse_project_filenames(project_file)
    actual_files_raw = scan_folder_filenames(folder_path)
    # 依 ignore 規則過濾
    project_files = {f for f in project_files_raw if not is_ignored(f)}
    actual_files  = {f for f in actual_files_raw  if not is_ignored(f)}
    matched = sorted(project_files & actual_files)
    missing = sorted(project_files - actual_files)
    extra = sorted(actual_files - project_files)
    return matched, missing, extra


# ----------------------------
# GUI 元件
# ----------------------------

# --- GUI 程式設計中的修改：.prproj 改為 .prproj / .aepx ---

# --- 行 1：專案檔案選擇（支援 prproj / aepx） ---
prproj_row = tk.Frame(root)
prproj_row.pack(pady=2, fill='x')

tk.Label(prproj_row, text="🎬 專案檔案（.prproj / .aepx）").pack(side='left')
prproj_entry = tk.Entry(prproj_row, width=70)
prproj_entry.pack(side='left', padx=5, expand=True, fill='x')
tk.Button(prproj_row, text="選擇檔案", command=lambda: browse_prproj()).pack(side='right')

# --- 行 2：素材資料夾
folder_row = tk.Frame(root)
folder_row.pack(pady=2, fill='x')

tk.Label(folder_row, text="📁 素材資料夾路徑").pack(side='left')
folder_entry = tk.Entry(folder_row, width=70)
folder_entry.pack(side='left', padx=5, expand=True, fill='x')
tk.Button(folder_row, text="選擇資料夾", command=lambda: browse_folder()).pack(side='right')

# --- 行 3：報告輸出資料夾 + 記住路徑
output_row = tk.Frame(root)
output_row.pack(pady=2, fill='x')

tk.Label(output_row, text="📂 報告輸出資料夾").pack(side='left')
output_entry = tk.Entry(output_row, width=60)
output_entry.pack(side='left', padx=5, expand=True, fill='x')
k_out_btn = tk.Button(output_row, text="選擇資料夾", command=lambda: browse_output_folder())
k_out_btn.pack(side='left', padx=4)

remember_output_var = tk.BooleanVar(value=False)
remember_output_chk = tk.Checkbutton(output_row, text="記住路徑", variable=remember_output_var)
remember_output_chk.pack(side='left')

# --- 行 4：LINE UID + 記住 UID
uid_row = tk.Frame(root)
uid_row.pack(pady=2, fill='x')

tk.Label(uid_row, text="👤 LINE UID").pack(side='left')
uid_entry = tk.Entry(uid_row, width=60, show="*")
uid_entry.pack(side='left', padx=5, expand=True, fill='x')

remember_uid_var = tk.BooleanVar(value=False)
remember_uid_chk = tk.Checkbutton(uid_row, text="記住 ID", variable=remember_uid_var)
remember_uid_chk.pack(side='left')

# --- 行 5：開始比對按鈕
start_btn = tk.Button(root, text="開始比對並產出報告", command=lambda: run_compare(), bg="lightblue")
start_btn.pack(pady=8)

# --- 行 6：輸出文字框
output_text = scrolledtext.ScrolledText(root, width=100, height=30)
output_text.pack(padx=10, pady=10, fill='both', expand=True)

# ----------------------------
# 與 GUI 互動的函式
# ----------------------------
# --- 行 1：專案檔案選擇 ---
def browse_prproj():
    path = filedialog.askopenfilename(filetypes=[
        ("所有支援的專案檔案", "{*.prproj} {*.aepx} {*.prproj.gz}")
    ])
    if path:
        prproj_entry.delete(0, tk.END)
        prproj_entry.insert(0, path)

# --- 行 2：素材資料夾 ---
def browse_folder():
    path = filedialog.askdirectory()
    if path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, path)

# --- 行 3：報告輸出資料夾 ---
def browse_output_folder():
    path = filedialog.askdirectory()
    if path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, path)

# 從 .prproj 路徑中提取專案名稱
def extract_project_label(prproj_path: str | Path) -> str:
    """
    回傳專案名稱（= '07_終極專案打包檔' 的上一層資料夾名稱）
    若不符合預期結構，則退而求其次回傳 .prproj 檔名（去掉副檔名）
    """
    p = Path(prproj_path)

    # (1) 正常情況：file → '07_終極專案打包檔' → <專案名稱>
    if len(p.parents) >= 2 and "終極專案打包檔" in p.parent.name:
        return p.parent.parent.name

    # (2) 萬一路徑少一層或命名不同，就回傳上一層
    if len(p.parents) >= 1:
        return p.parent.name

    # (3) 再不行就用檔名（不含副檔名）當 fallback
    return p.stem

# --- GUI 內警告提醒文字也一併修正 ---

def run_compare():
    prproj_path = prproj_entry.get().strip()
    folder_path = folder_entry.get().strip()
    output_path = output_entry.get().strip()
    uid = uid_entry.get().strip()

    if not prproj_path or not folder_path or not output_path or not uid:
        messagebox.showwarning("缺少輸入", "請確認已填入專案檔（.prproj / .aepx）、素材資料夾、輸出資料夾與 LINE UID")
        return
    
    # 先抓專案標籤，供後續 LINE 訊息與檔名使用
    project_label = extract_project_label(prproj_path)
    report_date = datetime.now().strftime("%Y-%m-%d")

    try:
        matched, missing, extra = compare_filenames(prproj_path, folder_path)
        output_text.delete(1.0, tk.END)
        #輸出專案名稱
        output_text.insert(tk.END, f"[{project_label}]\n✅ 對應成功的素材：{len(matched)}\n")
        for f in matched:
            output_text.insert(tk.END, f"  ✅ {f}\n")
        output_text.insert(tk.END, f"❌ 專案中使用但資料夾找不到：{len(missing)}\n")
        for f in missing:
            output_text.insert(tk.END, f"  ❌ {f}\n")
        output_text.insert(tk.END, f"⚠️ 資料夾中多餘素材（未在專案引用）：{len(extra)}\n")
        for f in extra:
            output_text.insert(tk.END, f"  ⚠️ {f}\n")

        # 報告內容組合
        report_date = datetime.now().strftime("%Y-%m-%d")
        text_report = f"""# Premiere 專案素材比對報告

## 專案檔案
- {prproj_path}

## 素材資料夾
- {folder_path}

---

## 對應成功的素材（共 {len(matched)} 筆）
""" + '\n'.join(f"- {f}" for f in matched) + f"""

---

## 專案中使用但資料夾找不到（共 {len(missing)} 筆）
""" + '\n'.join(f"- {f}" for f in missing) + f"""

---

## 資料夾中多餘素材（未在專案引用）（共 {len(extra)} 筆）
""" + '\n'.join(f"- {f}" for f in extra)

        # 寫入純文字檔案
        txt_file = Path(output_path) / f"Compare_report{project_label}_{report_date}_{report_date}.txt"
        txt_file.write_text(text_report, encoding="utf-8")

        messagebox.showinfo("報告完成", f"純文字報告儲存於：\n{txt_file}")

        # 傳送 LINE 精簡報告
        summary_text = f"""📊 [{project_label}]素材比對結果（{report_date})\n✅ 對應成功素材數量：{len(matched)}\n\n❌ 專案中使用但資料夾找不到素材（共 {len(missing)} 筆）：\n""" + '\n'.join(f"- {f}" for f in missing)
        send_to_lambda(summary_text[:4000], uid)

        # 最後儲存設定（若有勾選）
        save_config()

    except Exception as e:
        # 1) 將完整 traceback 寫入 log
        log_path = write_error_log(project_label, e)
        # 2) GUI 提示
        messagebox.showerror(
            "解析失敗",
            f"專案「{project_label}」解析失敗，已寫入日誌：\n{log_path}"
        )
        # 3) 傳 LINE 簡訊（只帶錯誤摘要）
        error_msg = (
            f"[{project_label}] ❗ 專案解析失敗！\n"
            f"🚨 錯誤：{e.__class__.__name__} - {e}\n"
            f"🪵 詳細日誌已寫入：{log_path.name}"
        )
        send_to_lambda(error_msg[:4000], uid)

# ----------------------------
# 啟動時載入偏好
# ----------------------------
config = load_config()
if config.get("line_uid"):
    uid_entry.insert(0, config["line_uid"])
    remember_uid_var.set(True)
if config.get("output_dir"):
    output_entry.insert(0, config["output_dir"])
    remember_output_var.set(True)

# ----------------------------
# 關閉視窗前儲存偏好
# ----------------------------

def on_quit():
    save_config()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_quit)

# ----------------------------
# 事件迴圈
# ----------------------------
if __name__ == "__main__":
    root.mainloop()
