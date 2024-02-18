import tkinter as tk
from tkinter import filedialog, ttk,messagebox
import dialog_functions as df
from excel_handler import ExcelHandler
import pandas as pd
import  os
def open_file(self):
    # 如果已经打开了一个文件，先关闭它
    if self.excel_handler is not None:
        self.excel_handler.close()
        # 移除所有的标签页
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

    self.file_path = filedialog.askopenfilename()
    if self.file_path:
        self.excel_handler = ExcelHandler(self.file_path)
        update_tables(self)


def update_tables(self):
    if not self.excel_handler:
        return
    if self.excel_handler.file_extension == ".csv":
        _update_table(self,"CSV", self.excel_handler.data)
    else:
        for sheet_name in self.excel_handler.get_sheet_names():
            sheet_data = self.excel_handler.data[sheet_name]
            _update_table(self,sheet_name, sheet_data)
            # 创建表格并添加到Notebook
            # tree = ttk.Treeview(self.notebook)
            # self.notebook.add(tree, text=sheet_name)
            # tree["columns"] = list(sheet_data.columns)
            # tree["show"] = "headings"
            # for col in sheet_data.columns:
            #     tree.heading(col, text=col)
            #
            # for index, row in sheet_data.iterrows():
            #     tree.insert("", "end", values=row.tolist())
def _update_table(self, name, data):
    tree = ttk.Treeview(self.notebook)
    self.notebook.add(tree, text=name)
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"
    for col in data.columns:
        tree.heading(col, text=col)

    for index, row in data.iterrows():
        tree.insert("", "end", values=row.tolist())
def get_selected_sheet_name(self):
    # 通过 notebook 组件获取当前选中的 tab 名字
    return self.notebook.tab(self.notebook.select(), "text")

# 修改列函数
def pj_column(self):
    # 弹出对话框
    self.dialog = tk.Toplevel(self)
    self.dialog.grab_set()  # 将对话框设置为模态
    self.dialog.title("Modify Column")
    prompt = "Tips:选择你要拼接的列，并在输入框中输入你想要拼接的内容。"  # 你想要显示的提示语
    lbl_prompt = tk.Label(self.dialog, text=prompt)
    lbl_prompt.pack()
    # 创建一个Listbox显示所有列
    self.column_listbox = tk.Listbox(self.dialog)
    # 获取当前选中的 sheet 名字
    sheet_name = get_selected_sheet_name(self)
    for column in self.excel_handler.get_columns(sheet_name):
        self.column_listbox.insert(tk.END, column)
    self.column_listbox.pack()
    # 创建一个 Label 显示文本
    self.input_label = tk.Label(self.dialog, text="该列前拼接:")
    self.input_label.pack()
    # 创建一个Entry获取用户的输入
    self.user_input1 = tk.Entry(self.dialog)
    self.user_input1.pack()
    self.input_label2 = tk.Label(self.dialog, text="该列后拼接:")
    self.input_label2.pack()
    # 创建一个Entry获取用户的输入
    self.user_input2 = tk.Entry(self.dialog)
    self.user_input2.pack()
    # 创建一个确定按钮
    self.confirm_button = tk.Button(self.dialog, text="Confirm", command=lambda: df.pj_confirm(self))
    self.confirm_button.pack()


def split_column(self):
    # 弹出对话框
    self.dialog = tk.Toplevel(self)
    self.dialog.grab_set()  # 将对话框设置为模态
    self.dialog.title("Modify Column")
    # 在这里添加提示语（Label）
    prompt = "Tips:选择你要拆分的列。"  # 你想要显示的提示语
    lbl_prompt = tk.Label(self.dialog, text=prompt)
    lbl_prompt.pack()
    # self.protocol("WM_DELETE_WINDOW", lambda: self.on_close())
    # 在主界面拦截关闭按钮点击事件
    # 创建一个Listbox显示所有列
    self.column_listbox = tk.Listbox(self.dialog)
    # 获取当前选中的 sheet 名字
    sheet_name = get_selected_sheet_name(self)
    for column in self.excel_handler.get_columns(sheet_name):
        self.column_listbox.insert(tk.END, column)
    self.column_listbox.pack()
    # 创建一个确定按钮
    self.confirm_button = tk.Button(self.dialog, text="Confirm", command=lambda: df.split_confirm(self))
    self.confirm_button.pack()

def map_column(self):
    # 创建对话框
    self.dialog = tk.Toplevel(self)
    self.dialog.grab_set()  # 将对话框设置为模态
    self.dialog.title("映射")
    # 在这里添加提示语（Label）
    prompt = "Tips:选择你根据映射的列。"
    self.lbl_prompt = tk.Label(self.dialog, text=prompt)
    self.lbl_prompt.pack()
    # 创建一个Listbox显示所有列
    self.column_listbox = tk.Listbox(self.dialog)
    # 获取当前选中的 sheet 名字
    selected_sheet1 = get_selected_sheet_name(self)
    for column in self.excel_handler.get_columns(selected_sheet1):
        self.column_listbox.insert(tk.END, column)
    self.column_listbox.pack()
    # 创建一个下一步按钮
    self.btn_next = tk.Button(self.dialog, text="下一步", command=lambda: df.show_sheets(self, selected_sheet1))
    self.btn_next.pack()
# 自动编号
def bh_column(self):
    # 弹出对话框
    self.dialog = tk.Toplevel(self)
    self.dialog.grab_set()  # 将对话框设置为模态
    self.dialog.title("Modify Column")
    prompt = "Tips:请选择你要自动编号的列。"  # 你想要显示的提示语
    lbl_prompt = tk.Label(self.dialog, text=prompt)
    lbl_prompt.pack()
    # 创建一个Listbox显示所有列
    self.column_listbox = tk.Listbox(self.dialog)
    # 获取当前选中的 sheet 名字
    sheet_name = get_selected_sheet_name(self)
    for column in self.excel_handler.get_columns(sheet_name):
        self.column_listbox.insert(tk.END, column)
    self.column_listbox.pack()
    # 创建一个确定按钮
    self.confirm_button = tk.Button(self.dialog, text="Confirm", command=lambda: df.bh_confirm(self))
    self.confirm_button.pack()


def execute_code(self):
    code = self.txt_code.get("1.0", "end-1c")
    file_name, file_extension = os.path.splitext(self.file_path)
    new_file_name = file_name + "处理后" + file_extension
    try:
        # 尝试执行代码
        exec(code)
        # 如果成功，显示成功消息
        messagebox.showinfo("Success", "Code executed successfully!"+"处理后的文件保存在"+new_file_name)
    except Exception as e:
        # 如果执行出错，捕获异常并显示错误消息
        messagebox.showerror("Error", f"An error occurred: {e}")
