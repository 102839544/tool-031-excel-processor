#!/usr/bin/env python3
"""
Excel批处理工具 - 批量合并/拆分/转换Excel文件
"""
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

class App:
    def __init__(self, root):
        self.root = root
        root.title("Excel批处理工具 v1.0")
        root.geometry("650x550")
        self.files = []
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg="#2e7d32", height=60)
        f.pack(fill="x")
        tk.Label(f, text="📊 Excel批处理工具", font=("Arial",16,"bold"),
                 fg="white", bg="#2e7d32").pack(pady=15)
        main = tk.Frame(self.root, padx=20, pady=15)
        main.pack(fill="both", expand=True)
        
        bf = tk.Frame(main)
        bf.pack(fill="x", pady=5)
        tk.Button(bf, text="添加Excel文件", command=self.add_files,
                  bg="#2e7d32", fg="white", padx=12).pack(side="left", padx=5)
        tk.Button(bf, text="清空列表", command=self.clear,
                  bg="#d9534f", fg="white", padx=12).pack(side="left", padx=5)
        
        self.lb = tk.Listbox(main, font=("Consolas",10), bg="#f8f9fa", height=12)
        self.lb.pack(fill="both", expand=True, pady=10)
        
        # 功能按钮
        ff = tk.Frame(main)
        ff.pack(fill="x", pady=10)
        tk.Button(ff, text="合并到一个Excel", command=self.merge_excel,
                  bg="#1976d2", fg="white", font=("Arial",10,"bold"),
                  padx=15, pady=5).pack(side="left", padx=5)
        tk.Button(ff, text="转为CSV", command=self.to_csv,
                  bg="#388e3c", fg="white", padx=15, pady=5).pack(side="left", padx=5)
        tk.Button(ff, text="提取所有Sheet名", command=self.list_sheets,
                  bg="#f57c00", fg="white", padx=15, pady=5).pack(side="left", padx=5)
        
        self.status = tk.Label(main, text="请添加Excel文件（支持 .xlsx/.xls）",
                              font=("Arial",10), fg="gray", anchor="w")
        self.status.pack(fill="x")
    
    def add_files(self):
        fs = filedialog.askopenfilenames(title="选择Excel文件",
             filetypes=[("Excel文件","*.xlsx *.xls")])
        for f in fs:
            if f not in self.files:
                self.files.append(f)
                self.lb.insert("end", Path(f).name)
        self.status.config(text=f"已添加 {len(self.files)} 个文件")
    
    def clear(self):
        self.files.clear()
        self.lb.delete(0, "end")
        self.status.config(text="列表已清空")
    
    def merge_excel(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加Excel文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        out = filedialog.asksaveasfilename(title="保存合并后的Excel",
             defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not out: return
        
        try:
            dfs = []
            for f in self.files:
                df = pd.read_excel(f)
                dfs.append(df)
            merged = pd.concat(dfs, ignore_index=True)
            merged.to_excel(out, index=False)
            messagebox.showinfo("完成", f"成功合并 {len(self.files)} 个文件！\n保存至：{out}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    
    def to_csv(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加Excel文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        out_dir = filedialog.askdirectory(title="选择输出目录")
        if not out_dir: return
        
        try:
            ok = 0
            for f in self.files:
                df = pd.read_excel(f)
                csv_path = str(Path(out_dir) / (Path(f).stem + ".csv"))
                df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                ok += 1
            messagebox.showinfo("完成", f"成功转换 {ok} 个文件为CSV！")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    
    def list_sheets(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加Excel文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        
        result = []
        for f in self.files:
            xl = pd.ExcelFile(f)
            result.append(f"\n{Path(f).name}:\n  " + "\n  ".join(xl.sheet_names))
        
        messagebox.showinfo("Sheet列表", "\n".join(result))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
