import tkinter as tk
import column_funcitons as cf


## 拼接对话框函数
def pj_confirm(self):
    # 获取用户选择的列和输入的文本
    # 获取 Listbox 的选中项
    if self.column_listbox.curselection():
        selected_column = self.column_listbox.get(self.column_listbox.curselection())
        # 获取 Entry 的输入
        user_input1 = self.user_input1.get()
        user_input2 = self.user_input2.get()
        # 获取当前选中的 sheet 名字
        sheet_name = cf.get_selected_sheet_name(self)
        # 调用 excel_handler 的方法来处理 Excel 文件
        self.excel_handler.modify_column(sheet_name, selected_column, user_input1, user_input2)
    else:
        tk.messagebox.showerror("Error", "No column selected")
    # selected_column = self.column_listbox.get(tk.ACTIVE)
    # user_text = self.user_input.get()
    # # 拼接到prompt语句中
    # prompt = f"Please modify the sheet '{self.notebook.select()}' column '{selected_column}' as '{user_text}'"
    # print(prompt)


# 对话框内的拆分按钮函数
def split_confirm(self):
    # 获取用户选择的列和输入的文本
    # 获取 Listbox 的选中项
    if self.column_listbox.curselection():
        selected_column = self.column_listbox.get(self.column_listbox.curselection())
        # 获取当前选中的 sheet 名字
        sheet_name = cf.get_selected_sheet_name(self)
        # 调用 excel_handler 的方法来处理 Excel 文件
        self.excel_handler.split_column(sheet_name, selected_column)
    else:
        tk.messagebox.showerror("Error", "No column selected")


# 映射列的对话框函数
def show_sheets(self, selected_sheet1):
    #在删除列表前先获取选择值
    if not self.column_listbox.curselection():
        selected_column1 = None
    else:
        selected_column1 = self.column_listbox.get(self.column_listbox.curselection())
    # 删除提示语和Listbox
    self.column_listbox.pack_forget()
    # 更新提示语
    new_prompt = "Tips:选择含有映射内容的sheet表。"
    self.lbl_prompt.config(text=new_prompt)
    # 创建一个新的Listbox显示所有的sheet表
    self.sheet_listbox = tk.Listbox(self.dialog)
    for sheet_name in self.excel_handler.get_sheet_names():
        self.sheet_listbox.insert(tk.END, sheet_name)
    self.sheet_listbox.pack()

    # 更新下一步按钮的回调函数
    self.btn_next.configure(command=lambda: next_step_after_sheets(self, selected_sheet1, selected_column1))


def next_step_after_sheets(self, selected_sheet1, selected_column1):
    #在删除前获取选择的sheet2表
    selected_sheet2 = self.sheet_listbox.get(self.sheet_listbox.curselection())
    # 删除和Listbox
    self.column_listbox.pack_forget()
    # 获取选中的sheet表名
    # 更新提示语
    new_prompt = "Tips:选择要根据哪一列进行映射。"
    self.lbl_prompt.config(text=new_prompt)
    # 创建一个新的Listbox显示选中的sheet表的所有列
    self.column_listbox = tk.Listbox(self.dialog)
    for column in self.excel_handler.get_columns(selected_sheet2):
        self.column_listbox.insert(tk.END, column)
    self.column_listbox.pack()
    # 更新下一步按钮的回调函数
    self.btn_next.configure(
        command=lambda: next_step2_after_sheets(self, selected_sheet1, selected_column1, selected_sheet2))


def next_step2_after_sheets(self, selected_sheet1, selected_column1, selected_sheet2):
    if not self.column_listbox.curselection():
        selected_column2 = None
    else:
        selected_column2 = self.column_listbox.get(self.column_listbox.curselection())
    # 更新提示语
    new_prompt = "Tips:选择要将哪一列函数映射到前一sheet表中。"
    self.lbl_prompt.config(text=new_prompt)
    self.btn_next.configure(
        command=lambda: next_step3_after_sheets(self, selected_sheet1, selected_column1, selected_sheet2,selected_column2))
def next_step3_after_sheets(self, selected_sheet1, selected_column1, selected_sheet2,selected_column2):
    if not self.column_listbox.curselection():
        selected_column3 = None
    else:
        selected_column3 = self.column_listbox.get(self.column_listbox.curselection())
    # # 更新提示语
    # new_prompt = "Tips:选择要将哪一列函数映射到前一sheet表中。"
    # self.lbl_prompt.config(text=new_prompt)
    self.btn_next.configure(
        command=self.excel_handler.map_column(selected_sheet1, selected_column1, selected_sheet2, selected_column2, selected_column3))
# 自动编号对话函数
def bh_confirm(self):
    # 获取用户选择的列和输入的文本
    # 获取 Listbox 的选中项
    if self.column_listbox.curselection():
        selected_column = self.column_listbox.get(self.column_listbox.curselection())
        sheet_name = cf.get_selected_sheet_name(self)
        # 调用 excel_handler 的方法来处理 Excel 文件
        self.excel_handler.bh_excel_column(sheet_name, selected_column)
    else:
        tk.messagebox.showerror("Error", "No column selected")
# def shake(dialog, x, count):
#     if count <= 0:
#         return
#     x_mult = 1 if count % 2 == 0 else -1
#     dialog.geometry(f"+{dialog.winfo_x() + 5 * x_mult}+{dialog.winfo_y()}")
#     dialog.after(50, shake, dialog, x * -1, count - 1)
# def on_main_window_closing(dialog):
#     shake(dialog, 5, 6)
