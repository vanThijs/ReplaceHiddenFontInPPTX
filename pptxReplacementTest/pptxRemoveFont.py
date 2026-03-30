import zipfile
import os
import shutil
import json
import tkinter as tk
from tkinter import filedialog, messagebox

# File to store the "memory"
CONFIG_FILE = "pptx_gui_config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {"last_search": "Noto Sans Symbols", "last_replace": "Verdana"}

def save_config(search_str, replace_str):
    with open(CONFIG_FILE, 'w') as f:
        json.dump({"last_search": search_str, "last_replace": replace_str}, f)

def run_replacement():
    input_pptx = entry_file.get()
    search_for = entry_search.get()
    replace_with = entry_replace.get()

    if not input_pptx or not os.path.exists(input_pptx):
        messagebox.showerror("Error", "Please select a valid .pptx file.")
        return

    # Process names
    base_name = os.path.splitext(input_pptx)[0]
    output_pptx = f"{base_name}_modified.pptx"
    zip_temp = "temp_archive.zip"
    
    total_replacements = 0
    shutil.copy2(input_pptx, zip_temp)
    
    try:
        with zipfile.ZipFile(zip_temp, 'r') as zin:
            with zipfile.ZipFile(output_pptx, 'w') as zout:
                for item in zin.infolist():
                    buffer = zin.read(item.filename)
                    if item.filename.endswith('.xml'):
                        content = buffer.decode('utf-8', errors='ignore')
                        count = content.count(search_for)
                        if count > 0:
                            content = content.replace(search_for, replace_with)
                            buffer = content.encode('utf-8')
                            total_replacements += count
                    zout.writestr(item, buffer)
        
        save_config(search_for, replace_with)
        messagebox.showinfo("Success", f"Replaced {total_replacements} occurrences.\nSaved as: {os.path.basename(output_pptx)}")
        
    except Exception as e:
        messagebox.showerror("Processing Error", str(e))
    finally:
        if os.path.exists(zip_temp):
            os.remove(zip_temp)

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if filename:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, filename)

# --- GUI Setup ---
root = tk.Tk()
root.title("PPTX String Replacer")
root.geometry("500x250")
root.resizable(False, False)

config = load_config()

# Layout
tk.Label(root, text="Select PPTX File:").pack(pady=(10, 0), padx=20, anchor="w")
frame_file = tk.Frame(root)
frame_file.pack(fill="x", padx=20)
entry_file = tk.Entry(frame_file)
entry_file.pack(side="left", expand=True, fill="x")
tk.Button(frame_file, text="Browse", command=browse_file).pack(side="right", padx=(5, 0))

tk.Label(root, text="Search for:").pack(pady=(10, 0), padx=20, anchor="w")
entry_search = tk.Entry(root)
entry_search.insert(0, config['last_search'])
entry_search.pack(fill="x", padx=20)

tk.Label(root, text="Replace with:").pack(pady=(10, 0), padx=20, anchor="w")
entry_replace = tk.Entry(root)
entry_replace.insert(0, config['last_replace'])
entry_replace.pack(fill="x", padx=20)

tk.Button(root, text="Run Replacement", command=run_replacement, bg="#4CAF50", fg="white", height=2).pack(pady=20)

root.mainloop()
