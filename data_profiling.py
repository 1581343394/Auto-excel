import pandas as pd
from ydata_profiling import ProfileReport
import os
import tkinter as tk
def explore_data(file_path):
    # 获取文件扩展名
    file_name,file_extension = os.path.splitext(file_path)
    # 分割文件名和扩展名
    # 根据文件扩展名读取数据
    if file_extension == ".xlsx":
        df = pd.read_excel(file_path)
    elif file_extension == ".csv":
        df = pd.read_csv(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")

    # 展示报告
    profile = ProfileReport(df, title='Pandas Profiling Report', explorative=True)
    # 在文件名后面添加 "副本"
    new_file_name = file_name + "数据探索报告" + '.html'
    # 将报告导出为HTML文件
    profile.to_file(new_file_name)
    tk.messagebox.showinfo("提示", "数据探索完成，报告保存在 " + new_file_name)