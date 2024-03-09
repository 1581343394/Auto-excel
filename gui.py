import tkinter as tk
from tkinter import filedialog, ttk,Text
from tkinter.scrolledtext import ScrolledText
import column_funcitons as cf
import  gpt_hander as gh
import data_profiling as daf
class Application(tk.Tk):
    # 提示语事件
    def on_txt_input_click(self, event):
        if self.txt_input.get("1.0", "end-1c") == "请在这里输入你的请求并点击发送!":
            self.txt_input.delete("1.0", "end")
    #导入代码
    def import_txt(self):
        # 打开文件对话框，导入文本文件
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as file:
                content = file.read()
                self.txt_code.delete("1.0", tk.END)
                self.txt_code.insert("1.0", content)
    #保存代码
    def save_as_txt(self):
        # 打开文件对话框，另存为文本文件
        file_path = filedialog.asksaveasfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            content = self.txt_code.get("1.0", tk.END)
            with open(file_path, 'w') as file:
                file.write(content)

    def on_model_selected(self, event):
        selected_model_type = self.selected_model_type.get()
        # 在这里根据选择的模型类型更新 config.ini 文件中的 selected_model 项

    def on_model_selected(self, event):
        selected_model_type = self.selected_model_type.get()
    def __init__(self):
        super().__init__()
        self.title("Excel Processor")
        self.geometry("800x600")
        self.excel_handler = None
        self.file_path=None
        self.dialog = None
        self.configure(bg='lightgrey') # 设置背景色为淡灰色

        # 创建菜单栏
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=lambda: cf.open_file(self))
        menubar.add_cascade(label="File", menu=filemenu)
        self.config(menu=menubar)

        # 创建按钮和文本框的frame
        frame = tk.Frame(self)
        frame.pack(fill=tk.X, padx=5, pady=5)

        # 创建拆分列按钮
        self.btn_split_column = tk.Button(frame, text="拆分列", command=lambda: cf.split_column(self))
        self.btn_split_column.pack(side="left", padx=5, pady=5)

        # 创建映射列按钮
        self.btn_map_column = tk.Button(frame, text="映射列", command=lambda: cf.map_column(self))
        self.btn_map_column.pack(side="left", padx=5, pady=5)

        # 添加修改列按钮
        self.modify_column_button = tk.Button(frame, text="拼接列内容", command=lambda: cf.pj_column(self))
        self.modify_column_button.pack(side="left", padx=5, pady=5)

        # 添加自动编号按钮
        self.modify_column_button = tk.Button(frame, text="自动编号",command=lambda: cf.bh_column(self))
        self.modify_column_button.pack(side="left", padx=5, pady=5)
        # 添加数据探索
        self.modify_column_button = tk.Button(frame, text="数据探索", command=lambda:daf.explore_data(self.file_path))
        self.modify_column_button.pack(side="left", padx=5, pady=5)
        # 创建Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")

        # 创建文本框
        frame_txt1 = ttk.Frame(self.notebook)
        frame_txt2 = ttk.Frame(self.notebook)
        self.notebook.add(frame_txt1, text='Input')
        self.notebook.add(frame_txt2, text='Code')

        self.txt_input = ScrolledText(frame_txt1, height=10)
        self.txt_input.pack(expand=True, fill="both", side="left", padx=5, pady=5)
        # 将提示语添加到文本框中
        self.txt_input.insert("1.0", "请在这里输入你的请求并点击发送!")
        # 绑定事件
        self.txt_input.bind("<FocusIn>", self.on_txt_input_click)
        # 创建左边的frame
        left_frame = tk.Frame(frame_txt2)
        left_frame.pack(side="left", fill="both", expand=True)

        # 创建右边的frame
        right_frame = tk.Frame(frame_txt2)
        right_frame.pack(side="right", fill="both", expand=True)

        # 创建左上的文本框
        self.txt_code = ScrolledText(left_frame, height=10)
        self.txt_code.pack(expand=True, fill="both", side="top", padx=5, pady=5)
        # 创建导入和另存为按钮的frame
        button_frame_txt2 = tk.Frame(left_frame)
        button_frame_txt2.pack(side="top", fill="x", padx=5, pady=5)

        # 创建导入按钮
        self.btn_import = tk.Button(button_frame_txt2, text="导入", command=self.import_txt)
        self.btn_import.pack(side="left", padx=5, pady=5)

        # 创建另存为按钮
        self.btn_save_as = tk.Button(button_frame_txt2, text="另存为", command=self.save_as_txt)
        self.btn_save_as.pack(side="left", padx=5, pady=5)
        # 创建左下的文本框
        # self.txt_pip = ScrolledText(left_frame, height=10)
        # self.txt_pip.pack(expand=True, fill="both", side="top", padx=5, pady=5)

        # 创建右边的文本框
        self.txt_explain = ScrolledText(right_frame, height=10)
        self.txt_explain.pack(expand=True, fill="both", side="top", padx=5, pady=5)
        # 创建下面按钮的frame
        btn_frame = tk.Frame(self)
        btn_frame.pack(side="bottom", padx=10, pady=10)

        # 创建发送按钮
        self.btn_send = tk.Button(btn_frame, text="发送", command=lambda: gh.GPTHandler.send_promt(self), bg='green',
                                  fg='white')  # 设置按钮颜色为绿色，字体颜色为白色
        self.btn_send.grid(row=0, column=1, padx=10, pady=10)

        # 创建执行按钮
        self.btn_execute = tk.Button(btn_frame, text="执行", command=lambda:cf.execute_code(self), bg='green',
                                     fg='white')  # 设置按钮颜色为绿色，字体颜色为白色
        self.btn_execute.grid(row=0, column=2, padx=10, pady=10)
        # 创建下拉框
        model_types = ["gpt-3.5", "文心一言3.0", "智普清言","gpt-4.0","claude3"]  # 你的模型类型列表
        self.selected_model_type = tk.StringVar(self)
        self.selected_model_type.set(model_types[0])  # 设置默认选择
        model_dropdown = ttk.Combobox(btn_frame, textvariable=self.selected_model_type, values=model_types)
        model_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        model_dropdown.bind("<<ComboboxSelected>>", self.on_model_selected)  # 绑定选择事件



