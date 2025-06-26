import pandas as pd
import glob
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl

def standardize_excel_openpyxl(input_path, output_path):
    try:
        wb = openpyxl.load_workbook(input_path)
        wb.save(output_path)
        return True
    except Exception as e:
        return False

def batch_standardize(folder, result_text):
    file_paths = glob.glob(os.path.join(folder, "*.xlsx"))
    file_paths = [f for f in file_paths if not os.path.basename(f).startswith('~$') and not os.path.basename(f).startswith('标准化_')]
    std_files = []
    failed_files = []
    for file in file_paths:
        std_file = os.path.join(folder, '标准化_' + os.path.basename(file))
        ok = standardize_excel_openpyxl(file, std_file)
        if ok:
            std_files.append(std_file)
            result_text.insert(tk.END, f"已标准化: {os.path.basename(file)}\n")
            result_text.see(tk.END)
        else:
            failed_files.append(file)
            result_text.insert(tk.END, f"标准化失败: {os.path.basename(file)}，请用Excel手动另存为标准xlsx再重试！\n")
            result_text.see(tk.END)
    return std_files

def merge_excels(folder, result_text):
    try:
        dt_str = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_path = os.path.join(folder, f"合并结果_{dt_str}.xlsx")
        # 只合并标准化后的文件
        file_paths = glob.glob(os.path.join(folder, "标准化_*.xlsx"))
        if not file_paths:
            result_text.insert(tk.END, "未找到任何标准化后的xlsx文件！\n")
            result_text.see(tk.END)
            return
        combined_df = None
        for i, file in enumerate(file_paths):
            try:
                with pd.ExcelFile(file, engine="openpyxl") as xls:
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                        if df.empty:
                            # 尝试用 header=None 读取
                            df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
                            if not df_raw.empty:
                                # 用第一行作为表头
                                df_raw.columns = df_raw.iloc[0]
                                df = df_raw[1:].reset_index(drop=True)
                        if df.empty:
                            continue  # 跳过空表
                        if combined_df is None:
                            combined_df = df
                        else:
                            combined_df = pd.concat([combined_df, df], ignore_index=True)
            except Exception as e:
                result_text.insert(tk.END, f"读取文件 {file} 时发生错误: {e}\n")
                result_text.see(tk.END)
                continue
        if combined_df is None or combined_df.empty:
            result_text.insert(tk.END, "所有文件都没有数据！\n")
            result_text.see(tk.END)
            return
        combined_df = combined_df.fillna("").astype(str)
        combined_df.to_excel(output_path, index=False)
        result_text.insert(tk.END, f"合并成功! \n 共合并 {len(file_paths)} 个文件，结果保存为: {output_path}\n")
        result_text.see(tk.END)
    except Exception as e:
        result_text.insert(tk.END, f"发生错误: {e}\n")
        result_text.see(tk.END)
        messagebox.showerror("错误", f"发生错误: {e}")
    finally:
        # 合并结束后删除所有标准化文件
        std_files = glob.glob(os.path.join(folder, "标准化_*.xlsx"))
        for f in std_files:
            try:
                os.remove(f)
            except Exception as e:
                result_text.insert(tk.END, f"删除临时文件失败: {f}, 错误: {e}\n")
                result_text.see(tk.END)

def start_merge(result_text, folder_var):
    folder = folder_var.get()
    if not os.path.isdir(folder):
        messagebox.showerror("错误", "请选择有效的文件夹！")
        return
    result_text.insert(tk.END, "开始标准化所有Excel文件...\n")
    result_text.see(tk.END)
    std_files = batch_standardize(folder, result_text)
    if not std_files:
        result_text.insert(tk.END, "没有可用的标准化文件，终止合并。\n")
        result_text.see(tk.END)
        return
    result_text.insert(tk.END, "标准化完成，开始合并...\n")
    result_text.see(tk.END)
    merge_excels(folder, result_text)

def select_folder(folder_var, result_text):
    folder = filedialog.askdirectory()
    if folder:
        folder_var.set(folder)

def main():
    root = tk.Tk()
    root.title("Excel合并工具")
    folder_var = tk.StringVar(value=os.path.join(os.getcwd(), "excel"))
    tk.Label(root, text="选择Excel文件夹:").pack()
    folder_entry = tk.Entry(root, textvariable=folder_var, width=50)
    folder_entry.pack()
    result_text = scrolledtext.ScrolledText(root, width=60, height=10)
    result_text.pack()
    tk.Button(root, text="浏览", command=lambda: select_folder(folder_var, result_text)).pack()
    tk.Button(root, text="开始合并", command=lambda: start_merge(result_text, folder_var)).pack()
    usage = (
        "使用说明：\n"
        "1. 将所有需要合并的xlsx文件放到同一个文件夹下。\n"
        "2. 选择该文件夹。\n"
        "3. 点击 开始合并，合并结果会生成在同一文件夹下，文件名为 合并结果_日期时间.xlsx。\n"
    )
    tk.Label(root, text=usage, justify="left", fg="blue").pack()
    root.mainloop()

if __name__ == "__main__":
    main()
