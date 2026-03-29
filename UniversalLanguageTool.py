import pandas as pd
from openai import OpenAI
import time
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
from tkinterdnd2 import DND_FILES, TkinterDnD #拖拽上传文件

# ================= 配置文件路径 =================
CONFIG_FILE = "translator_config.json"

# ================= 主流大模型预设字典 =================
# 格式：{"模型名称": "对应的接口地址"}
DEFAULT_PRESETS = {
    # 🌍 --- 国外顶尖模型 (使用需科学上网环境) ---
    "gpt-4o": "https://api.openai.com/v1",
    "gemini-2.5-flash": "https://generativelanguage.googleapis.com/v1beta/openai/",
    "llama-3.3-70b-versatile": "https://api.groq.com/openai/v1",
    "mistral-large-latest": "https://api.mistral.ai/v1",
    "meta-llama/Llama-3-70b-chat-hf": "https://api.together.xyz/v1",
    
    # 🇨🇳 --- 国内顶尖模型 (国内网络直连) ---
    "deepseek-chat": "https://api.deepseek.com",
    "qwen-plus": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "moonshot-v1-8k": "https://api.moonshot.cn/v1",
    "glm-4-flash": "https://open.bigmodel.cn/api/paas/v4",
    "deepseek-ai/DeepSeek-V3": "https://api.siliconflow.cn/v1",
    "Baichuan4": "https://api.baichuan-ai.com/v1"
}

API_PRESETS = DEFAULT_PRESETS.copy()

# ================= 配置保存与读取逻辑 =================
def load_config():
    """软件启动时读取本地配置"""
    global API_PRESETS
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
                # 动态加载用户自定义保存的模型预设
                if "custom_presets" in config:
                    API_PRESETS.update(config["custom_presets"])
                    combo_model['values'] = list(API_PRESETS.keys())
                
                if "api_key" in config and config["api_key"]:
                    entry_key.delete(0, tk.END)
                    entry_key.insert(0, config["api_key"])
                if "model" in config and config["model"]:
                    combo_model.set(config["model"])
                if "base_url" in config and config["base_url"]:
                    entry_baseurl.delete(0, tk.END)
                    entry_baseurl.insert(0, config["base_url"])
                if "source_lang" in config and config["source_lang"]:
                    entry_source.delete(0, tk.END)
                    entry_source.insert(0, config["source_lang"])
                if "target_lang" in config and config["target_lang"]:
                    entry_target.delete(0, tk.END)
                    entry_target.insert(0, config["target_lang"])
        except Exception as e:
            print(f"读取配置失败: {e}")

def save_config():
    """点击翻译时自动保存当前配置到本地，包括用户手写的新模型"""
    current_model = combo_model.get().strip()
    current_url = entry_baseurl.get().strip()
    
    # 读取旧配置文件中的自定义预设（如果有的话）
    custom_presets = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                old_config = json.load(f)
                custom_presets = old_config.get("custom_presets", {})
        except:
            pass
            
    # 如果用户输入了一个有效的模型和地址，并且它不在默认字典里（或者地址被修改了），就存入自定义预设
    if current_model and current_url:
        if current_model not in DEFAULT_PRESETS or DEFAULT_PRESETS[current_model] != current_url:
            custom_presets[current_model] = current_url
            API_PRESETS[current_model] = current_url
            combo_model['values'] = list(API_PRESETS.keys()) # 实时更新下拉菜单

    config = {
        "api_key": entry_key.get().strip(),
        "model": current_model,
        "base_url": current_url,
        "source_lang": entry_source.get().strip(),
        "target_lang": entry_target.get().strip(),
        "custom_presets": custom_presets # 把自定义模型写入账本
    }
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"保存配置失败: {e}")

# ================= 核心翻译逻辑 =================
def translate_text(client, text, source_lang, target_lang, model_name):
    """调用大模型 API 翻译单条文本"""
    if pd.isna(text) or str(text).strip() == "":
        return ""
    
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {
                    "role": "system", 
                    "content": (
                        f"你是一个专业的游戏本地化翻译专家。\n"
                        f"任务：将给定的{source_lang}文本翻译成{target_lang}。\n"
                        "【核心强制规则】：\n"
                        "1. 占位符保护：严禁翻译或修改任何大括号内的内容（如 {0}, {name} 等），必须原样保留。\n"
                        "2. 标签保护：严禁翻译或修改任何尖括号标签（如 <color=#xxxx>, </color>, <size=20> 等），必须原样保留其位置和字符。\n"
                        "3. 标签内容翻译：标签包裹的中文文本（例如 <color>内容</color> 中的“内容”）必须准确翻译成目标语言。\n"
                        "4. 格式要求：严禁在翻译结果中添加任何额外的解释、引号或注音。\n"
                        "5. 语境：保持游戏UI文本的简洁、严谨。"
                    )
                },
                {"role": "user", "content": str(text)},
            ],
            temperature=0.1,
            stream=False
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"[翻译报错: {e}]"

# ================= 界面交互逻辑 =================
def start_translation():
    save_config()

    api_key = entry_key.get().strip()
    model_name = combo_model.get().strip()
    base_url = entry_baseurl.get().strip()
    input_file = entry_file.get().strip()
    s_lang = entry_source.get().strip()
    
    t_langs_raw = entry_target.get().strip().replace('，', ',')
    t_langs = [lang.strip() for lang in t_langs_raw.split(',') if lang.strip()]
    
    target_range_raw = entry_target_col.get().strip().upper().replace('，', ',')
    range_parts = re.split(r'[,\-~]', target_range_raw) 
    range_parts = [p.strip() for p in range_parts if p.strip()]
    
    out_dir = entry_out_dir.get().strip()
    out_name = entry_out_name.get().strip()

    if not api_key or not input_file or not model_name or not range_parts or not t_langs:
        messagebox.showerror("参数缺失", "API Key、模型名称、Excel文件、翻译范围和目标语言为必填项！")
        return

    start_cell = range_parts[0]
    try:
        match_start = re.match(r"^([A-Z]+)(\d+)$", start_cell)
        if not match_start:
            raise ValueError("起始单元格格式错误，请输入如 A3 的格式！")
        
        col_letters = match_start.group(1)
        start_row_idx = int(match_start.group(2)) - 1 
        
        col_idx = 0
        for char in col_letters:
            col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
        col_idx -= 1
        
    except Exception as e:
        messagebox.showerror("格式错误", str(e))
        return

    btn_start.config(state=tk.DISABLED, text="正在翻译中...")
    progress_var.set(0)
    lbl_progress.config(text="进度: 0%")
    
    client_kwargs = {"api_key": api_key}
    if base_url:
        client_kwargs["base_url"] = base_url
    
    try:
        client = OpenAI(**client_kwargs)
        df = pd.read_excel(input_file, header=None)
        
        if col_idx >= len(df.columns) or start_row_idx >= len(df):
            raise ValueError("指定的起始单元格超出了当前Excel表格的数据范围！")

        end_row_idx = len(df) 
        if len(range_parts) > 1:
            match_end = re.match(r"^([A-Z]*)(\d+)$", range_parts[1])
            if not match_end:
                raise ValueError("终止位置格式错误，请输入如 A16 的格式！")
            
            parsed_end_row = int(match_end.group(2))
            end_row_idx = min(parsed_end_row, len(df)) 
            
            if end_row_idx <= start_row_idx:
                raise ValueError("终止行号必须大于起始行号！")

        for i, lang in enumerate(t_langs):
            new_col_name = f"Temp_Result_{lang}"
            df.insert(col_idx + 1 + i, new_col_name, "")
            if start_row_idx > 0:
                df.iloc[start_row_idx - 1, col_idx + 1 + i] = f"{lang}翻译"

        total_rows_to_process = end_row_idx - start_row_idx
        total_tasks = total_rows_to_process * len(t_langs)
        progress_bar["maximum"] = total_tasks
        
        log_box.insert(tk.END, f"🚀 开始任务: {s_lang} -> {', '.join(t_langs)}\n")
        if len(range_parts) > 1:
            log_box.insert(tk.END, f"📍 范围限制: 第 {start_row_idx+1} 行 至 第 {end_row_idx} 行\n")
        log_box.see(tk.END)
        root.update()
        
        task_count = 0
        
        for r in range(start_row_idx, end_row_idx):
            raw_val = df.iloc[r, col_idx]
            
            if pd.isna(raw_val) or str(raw_val).strip() == "" or str(raw_val).strip().lower() == "nan":
                task_count += len(t_langs)
                progress_var.set(task_count)
                continue
                
            source_text = str(raw_val).strip()
            
            for i, lang in enumerate(t_langs):
                res_col_idx = col_idx + 1 + i
                existing_val = str(df.iloc[r, res_col_idx])
                
                if existing_val.strip() != "" and "[翻译报错" not in existing_val:
                    pass 
                else:
                    translated = translate_text(client, source_text, s_lang, lang, model_name)
                    df.iloc[r, res_col_idx] = translated
                    
                    display_source = source_text if len(source_text) < 10 else source_text[:8] + "..."
                    log_box.insert(tk.END, f"[第{r+1}行] {display_source} -> [{lang}] 完成\n")
                    log_box.see(tk.END) 
                
                task_count += 1
                progress_var.set(task_count)
                percent = int((task_count / total_tasks) * 100)
                lbl_progress.config(text=f"进度: {percent}% ({task_count}/{total_tasks})")
                root.update() 
                time.sleep(0.1)

        if not out_dir:
            out_dir = os.path.dirname(input_file)
            
        if not out_name:
            out_name = f"多语言处理完成_{int(time.time())}.xlsx"
        elif not out_name.endswith(('.xlsx', '.xls')):
            out_name += ".xlsx"
            
        output_file = os.path.join(out_dir, out_name)
        df.to_excel(output_file, index=False, header=False)
        log_box.insert(tk.END, f"\n🎉 恭喜，全部翻译完成！\n")
        log_box.see(tk.END)
        open_folder = messagebox.askyesno("任务成功", f"翻译完成！文件已保存为:\n{output_file}\n\n是否立即打开所在文件夹？")
        if open_folder:
            target_dir = os.path.dirname(output_file)
            if os.name == 'nt':
                os.startfile(target_dir)

    except Exception as e:
        messagebox.showerror("运行发生错误", f"详细报错信息:\n{str(e)}")
    finally:
        btn_start.config(state=tk.NORMAL, text="🚀 开始执行翻译")

def select_file():
    path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
    if path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, path)

def select_out_dir():
    path = filedialog.askdirectory()
    if path:
        entry_out_dir.delete(0, tk.END)
        entry_out_dir.insert(0, path)


# ================= 全局右键菜单功能 =================
def add_context_menu(root):
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="复制 (Copy)", command=lambda: root.focus_get().event_generate("<<Copy>>"))
    menu.add_command(label="粘贴 (Paste)", command=lambda: root.focus_get().event_generate("<<Paste>>"))
    menu.add_command(label="剪切 (Cut)", command=lambda: root.focus_get().event_generate("<<Cut>>"))
    menu.add_separator()
    
    def select_all():
        widget = root.focus_get()
        if isinstance(widget, (tk.Entry, ttk.Combobox)):
            widget.select_range(0, tk.END)
            widget.icursor(tk.END)
        elif isinstance(widget, tk.Text):
            widget.tag_add("sel", "1.0", "end")
            widget.mark_set("insert", "1.0")
            widget.see("insert")
            
    menu.add_command(label="全选 (Select All)", command=select_all)

    def show_menu(event):
        widget = event.widget
        if isinstance(widget, (tk.Entry, tk.Text, ttk.Combobox)):
            widget.focus_set()
            menu.tk_popup(event.x_root, event.y_root)

    root.bind_all("<Button-3>", show_menu)


# ================= 窗口 UI 布局 =================
root = TkinterDnD.Tk() 

root.title("多语言处理工具 V1.0.0")
root.geometry("850x700")

add_context_menu(root)

main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=1)

canvas = tk.Canvas(main_frame, highlightthickness=0, borderwidth=0)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

scrollable_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
scrollable_frame.bind("<Configure>", on_frame_configure)

def _on_mousewheel(event):
    if event.widget == log_box:
        return
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
canvas.bind_all("<MouseWheel>", _on_mousewheel)

def on_canvas_configure(event):
    canvas.itemconfig(canvas.find_withtag("all")[0], width=event.width)
canvas.bind("<Configure>", on_canvas_configure)

# --- 1. API 配置区 ---
frame_api = tk.LabelFrame(scrollable_frame, text="1. 大模型 API 配置", font=("Arial", 10, "bold"), padx=15, pady=10)
frame_api.pack(fill=tk.X, padx=20, pady=(15, 10))

tk.Label(frame_api, text="API Key (将自动保存在本地):").pack()
entry_key = tk.Entry(frame_api, width=72, show="*") 
entry_key.pack()

# 👉 核心修改区：上下位置调换
tk.Label(frame_api, text="模型名称 (Model - 可下拉选择预设，或手动输入):").pack(pady=(5, 0))
combo_model = ttk.Combobox(frame_api, width=70, values=list(API_PRESETS.keys()))
combo_model.insert(0, "deepseek-chat") 
combo_model.pack()

tk.Label(frame_api, text="Base URL (接口地址 - 随模型变化，也可手动修改):").pack(pady=(5, 0))
entry_baseurl = tk.Entry(frame_api, width=72)
entry_baseurl.insert(0, "https://api.deepseek.com") 
entry_baseurl.pack()

# 💡 联动事件 1：选中下拉框模型时，自动填入对应的地址
def on_model_select(event):
    selected_model = combo_model.get().strip()
    if selected_model in API_PRESETS:
        entry_baseurl.delete(0, tk.END)
        entry_baseurl.insert(0, API_PRESETS[selected_model])

combo_model.bind("<<ComboboxSelected>>", on_model_select)

# 💡 联动事件 2：一旦用户手动修改了地址，且与当前预设模型不符，自动清空模型框让你自己写
def on_baseurl_edit(event):
    current_url = entry_baseurl.get().strip()
    current_model = combo_model.get().strip()
    
    if current_model in API_PRESETS and API_PRESETS[current_model] != current_url:
        combo_model.set("") # 清空模型框

entry_baseurl.bind("<KeyRelease>", on_baseurl_edit)


# --- 左右并排容器区 ---
container_middle = tk.Frame(scrollable_frame)
container_middle.pack(fill=tk.X, padx=20, pady=10)

# --- 2. 翻译任务配置区 ---
frame_input = tk.LabelFrame(container_middle, text="2. 翻译任务配置", font=("Arial", 10, "bold"), padx=15, pady=10)
frame_input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

# 修改一下这里的 Label 提示语，告诉用户可以拖拽
tk.Label(frame_input, text="选择输入 Excel 表格 (支持直接拖入文件):").pack(anchor="w")
frame_file_inner = tk.Frame(frame_input)
frame_file_inner.pack(fill=tk.X)
entry_file = tk.Entry(frame_file_inner)
entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
tk.Button(frame_file_inner, text="浏览", command=select_file).pack(side=tk.RIGHT)

# 👇 从这里开始新增拖拽功能的核心代码
def on_drop(event):
    # 获取拖入的文件路径 (Windows 环境下如果路径有空格，首尾会自动加上大括号，需要剥离)
    file_path = event.data
    if file_path.startswith('{') and file_path.endswith('}'):
        file_path = file_path[1:-1]
        
    # 简单校验一下是不是 Excel 文件
    if file_path.lower().endswith(('.xlsx', '.xls')):
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
    else:
        messagebox.showwarning("格式提醒", "请拖入 Excel 文件 (.xlsx 或 .xls) ")

# 把输入框注册为“可以接受文件拖入”的区域
entry_file.drop_target_register(DND_FILES)
entry_file.dnd_bind('<<Drop>>', on_drop)
# 👆 新增拖拽代码结束

tk.Label(frame_input, text="翻译范围 (如 A3 或 A3,A16):").pack(anchor="w", pady=(10, 0))
entry_target_col = tk.Entry(frame_input, width=18)
entry_target_col.insert(0, "A2") 
entry_target_col.pack(anchor="w")

tk.Label(frame_input, text="源语言:").pack(anchor="w", pady=(10, 0))
entry_source = tk.Entry(frame_input, width=15)
entry_source.insert(0, "中文")
entry_source.pack(anchor="w")

tk.Label(frame_input, text="目标语言(逗号分隔):").pack(anchor="w", pady=(10, 0))
entry_target = tk.Entry(frame_input, width=30)
entry_target.insert(0, "英语,日语")
entry_target.pack(anchor="w")

# --- 3. 输出文件配置区 ---
frame_output = tk.LabelFrame(container_middle, text="3. 输出文件配置", font=("Arial", 10, "bold"), padx=15, pady=10)
frame_output.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))

tk.Label(frame_output, text="输出保存文件夹 (不填则默认同目录):").pack(anchor="w")
frame_outdir_inner = tk.Frame(frame_output)
frame_outdir_inner.pack(fill=tk.X)
entry_out_dir = tk.Entry(frame_outdir_inner)
entry_out_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
tk.Button(frame_outdir_inner, text="选择", command=select_out_dir).pack(side=tk.RIGHT)

tk.Label(frame_output, text="自定义输出文件名 (例：CN_to_JP):").pack(anchor="w", pady=(10, 0))
entry_out_name = tk.Entry(frame_output, width=30)
entry_out_name.pack(anchor="w")

# --- 4. 执行与日志区 ---
frame_bottom = tk.Frame(scrollable_frame)
frame_bottom.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

btn_start = tk.Button(frame_bottom, text="🚀 开始执行翻译", command=start_translation, bg="#28a745", fg="white", font=("Arial", 12, "bold"), height=2, width=30)
btn_start.pack(pady=(10, 15))

lbl_progress = tk.Label(frame_bottom, text="进度: 0%", font=("Arial", 9))
lbl_progress.pack(side=tk.TOP)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(frame_bottom, variable=progress_var, maximum=100)
progress_bar.pack(fill=tk.X, expand=True, pady=(0, 10))

tk.Label(frame_bottom, text="运行日志:").pack(anchor="w")
log_box = tk.Text(frame_bottom, height=12, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9)) 
log_box.pack(fill=tk.BOTH, expand=True, pady=(5, 20))

load_config()

root.mainloop()