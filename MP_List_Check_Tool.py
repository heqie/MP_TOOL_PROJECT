#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT 
@File    ：MP List Check Tool.py
@Author  ：zxh
@Date    ：2025/3/19 16:20 
'''
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, scrolledtext
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import matplotlib.animation as animation
from read_excel import *
from check import *
from analyze import *

pd.options.mode.chained_assignment = None  # 禁用警告
# 解决中文显示问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

multiple = 1.2


class CheckFrame(tk.Frame):
    """
    Check界面
    """

    def __init__(self, master=None):
        """
        self值初始化
        :param master: NONE
        """
        super().__init__(master)
        self.master = master
        self.file_path = None
        self.df = None
        self.valid_glass_models = None
        self.valid_TOOL_version = None
        self.valid_IC = None
        self.checks_visible = False  # 用于跟踪功能按钮的显示状态
        self.create_widgets()
        self.sequence_error = None

    def create_widgets(self):
        """
        绘制UI界面
        :return:
        """
        # 文件路径显示区域
        self.label_file_path = tk.Label(self, text="文 件 路 径:")
        self.label_file_path.place(x=10, y=10)

        self.text_file_path = tk.Entry(self, state="readonly")
        self.text_file_path.place(x=90, y=10, width=320 * multiple + 20, height=25)

        # 文件选择按钮
        self.btn_open = tk.Button(self, text="...", width=3, bg='Gainsboro', command=self.open_file)
        self.btn_open.place(x=420 * multiple, y=10)

        # 重新读取文件按钮
        self.reload_var = tk.BooleanVar(value=True)  # 默认勾选
        self.chk_reload = tk.Checkbutton(self, text="reload excel", width=15, variable=self.reload_var)
        self.chk_reload.place(x=460 * multiple, y=10)

        # self.btn_reload_excel = tk.Button(self, text="重新读取文件", height=1, width=15, command=self.reload_excel,
        #                                   state=tk.DISABLED)
        # self.btn_reload_excel.place(x=90, y=50)

        # 一键检查按钮
        self.btn_check_all = tk.Button(self, text="一键检查", height=1, width=15, command=self.run_check_all,
                                       state=tk.DISABLED)
        self.btn_check_all.place(x=80, y=50)

        # 分项检查按钮
        self.btn_show_checks = tk.Button(self, text="分项检查", height=1, width=15, command=self.show_checks,
                                         state=tk.DISABLED)
        self.btn_show_checks.place(x=250 * multiple, y=50)

        # 序号校正按钮
        self.btn_sequence_revise = tk.Button(self, text="序号校正", height=1, width=15,
                                             command=lambda: self.fix_sequence_errors(self.sequence_error),
                                             state=tk.DISABLED)
        self.btn_sequence_revise.place(x=420 * multiple, y=50)

        # 检查按钮（初始状态为隐藏和禁用）
        self.btn_sequence = tk.Button(self, text="序列检查", height=1, width=15,
                                      command=lambda: self.run_check("序列检查", lambda: check_sequence(self.df)))
        self.btn_grade = tk.Button(self, text="等级检查", height=1, width=15,
                                   command=lambda: self.run_check("等级检查", lambda: check_grade(self.df.iloc[:, 1])))
        self.btn_ic_model = tk.Button(self, text="IC型号检查", height=1, width=15, command=lambda: self.run_check("IC型号检查",
                                                                                                              lambda: check_ic_model(
                                                                                                                  self.df.iloc[
                                                                                                                  :, 6],
                                                                                                                  self.valid_IC)))
        self.btn_glass_model = tk.Button(self, text="玻璃型号检查", height=1, width=15,
                                         command=lambda: self.run_check("玻璃型号检查",
                                                                        lambda: check_glass_model(self.df.iloc[:, 7],
                                                                                                  self.valid_glass_models)))
        self.btn_interface = tk.Button(self, text="接口定义检查", height=1, width=15, command=lambda: self.run_check("接口定义检查",
                                                                                                               lambda: check_interface(
                                                                                                                   self.df.iloc[
                                                                                                                   :,
                                                                                                                   8])))
        self.btn_flash = tk.Button(self, text="flash检查", height=1, width=15, command=lambda: self.run_check("flash检查",
                                                                                                            lambda: check_flash(
                                                                                                                self.df.iloc[
                                                                                                                :, 9],
                                                                                                                self.df.iloc[
                                                                                                                :, 8])))
        self.btn_tool_version = tk.Button(self, text="TOOL版本检查", height=1, width=15,
                                          command=lambda: self.run_check("TOOL版本检查", lambda: check_tool_version(self.df,
                                                                                                                self.valid_TOOL_version)))
        self.btn_date_format = tk.Button(self, text="发布日期检查", height=1, width=15,
                                         command=lambda: self.run_check("发布日期检查", lambda: check_date_format(self.df)))

        # 结果显示区域
        self.result_area = scrolledtext.ScrolledText(self)
        self.result_area.place(x=10 * multiple, y=90, width=590 * multiple, height=480 * multiple)
        self.result_area.config(state=tk.DISABLED)

        # 设置颜色标签
        self.result_area.tag_config("green", foreground="green")
        self.result_area.tag_config("red", foreground="red")

    def open_file(self):
        """
        打开文件夹并且写入文件路径，用于打开文件操作
        :return:
        """
        # 打开文件选择对话框
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        # 选择文件路径后，清一次屏
        self.result_area.config(state=tk.NORMAL)
        self.result_area.delete(1.0, tk.END)
        self.result_area.config(state=tk.DISABLED)
        if self.file_path:
            # 更新文件路径
            self.text_file_path.config(state="normal")
            self.text_file_path.delete(0, tk.END)
            self.text_file_path.insert(0, self.file_path)
            self.text_file_path.config(state="readonly")

            # 启用一键检查和分项检查按钮
            self.btn_check_all.config(state=tk.NORMAL)
            self.btn_show_checks.config(state=tk.NORMAL)

    def reload_excel(self):
        """
        加载文件，用于重新加载文件按键操作
        :return:
        """
        try:
            self.df = read_excel_with_merged_cells(self.file_path, sheet_name='MP Project List（internal）')
            if self.df is not None:
                # 读取对比文件
                valid_data_file = 'compare_information.xlsx'
                try:
                    valid_data = read_compare_excel(valid_data_file, sheet_names=None)
                    if valid_data is not None and not valid_data.empty:
                        self.valid_TOOL_version = valid_data.iloc[:, 1].dropna().tolist()
                        self.valid_glass_models = valid_data.iloc[:, 2].dropna().tolist()
                        self.valid_IC = valid_data.iloc[:, 0].dropna().tolist()
                    else:
                        self.valid_TOOL_version = []
                        self.valid_glass_models = []
                        self.valid_IC = []
                        messagebox.showwarning("警告", "对比文件为空或无效")
                except Exception as e:
                    self.valid_TOOL_version = []
                    self.valid_glass_models = []
                    self.valid_IC = []
                    messagebox.showerror("错误", f"读取对比文件失败: {e}")
            else:
                messagebox.showerror("错误", "读取Excel文件失败")
        except Exception as e:
            self.df = None
            messagebox.showerror("错误", f"读取Excel文件失败: {e}")

    def show_checks(self):
        """
        隐藏/显示分项检查功能按键
        :return:
        """
        if self.checks_visible:
            # 隐藏功能按钮
            self.result_area.place(x=10 * multiple, y=90, width=590 * multiple, height=480 * multiple)
            self.btn_sequence.place_forget()
            self.btn_grade.place_forget()
            self.btn_ic_model.place_forget()
            self.btn_glass_model.place_forget()
            self.btn_interface.place_forget()
            self.btn_flash.place_forget()
            self.btn_tool_version.place_forget()
            self.btn_date_format.place_forget()
            self.checks_visible = False
        else:
            # 显示功能按钮
            self.result_area.place(x=10 * multiple, y=170, width=590 * multiple, height=400 * multiple)
            self.btn_sequence.place(x=10 * multiple, y=90)
            self.btn_grade.place(x=160 * multiple, y=90)
            self.btn_ic_model.place(x=310 * multiple, y=90)
            self.btn_glass_model.place(x=465 * multiple, y=90)
            self.btn_interface.place(x=10 * multiple, y=130)
            self.btn_flash.place(x=160 * multiple, y=130)
            self.btn_tool_version.place(x=310 * multiple, y=130)
            self.btn_date_format.place(x=465 * multiple, y=130)
            self.checks_visible = True

    def run_check_all(self):
        """
        一次性检查所有项，用于一键检查按键操作
        :return:
        """
        if self.reload_var.get():  # 如果勾选框被勾选，默认重新加载文件
            self.reload_excel()
        checks = [
            ("序列检查", check_sequence(self.df)),
            ("等级检查", check_grade(self.df.iloc[:, 1])),
            ("IC型号检查", check_ic_model(self.df.iloc[:, 6], self.valid_IC)),
            ("玻璃型号检查", check_glass_model(self.df.iloc[:, 7], self.valid_glass_models)),
            ("接口定义检查", check_interface(self.df.iloc[:, 8])),
            ("flash有无检查", check_flash(self.df.iloc[:, 9], self.df.iloc[:, 8])),
            ("Tool版本号检查", check_tool_version(self.df, self.valid_TOOL_version)),
            ("发布日期检查", check_date_format(self.df)),
        ]
        # 清空结果显示区域
        self.result_area.config(state=tk.NORMAL)
        self.result_area.delete(1.0, tk.END)

        # 显示检查结果
        for check_name, errors in checks:
            self.result_area.insert(tk.END, f"{check_name}:\n")
            if not errors:
                self.result_area.insert(tk.END, "  PASS\n", "green")
                if check_name == "序列检查":
                    self.btn_sequence_revise.config(state=tk.DISABLED)
            else:
                self.result_area.insert(tk.END, "FAIL:\n", "red")
                self.result_area.insert(tk.END, "错误位置:\n", "red")
                for error in errors:
                    self.result_area.insert(tk.END, f" 第 {error[0]} 行, 错误值: {error[2]}\n", "red")

                # 序列检查有问题则启用序号校正按钮
                if check_name == "序列检查":
                    self.btn_sequence_revise.config(state=tk.NORMAL)
                    self.sequence_error = check_sequence_for_fix(self.df)

            self.result_area.insert(tk.END, "\n")

        self.result_area.config(state=tk.DISABLED)

    def run_check(self, check_name, check_func):
        """
        用于分项检查按键操作
        :param check_name: check项目名
        :param check_func: 对应check函数
        :return:
        """
        if self.reload_var.get():  # 如果勾选框被勾选，默认重新加载文件
            self.reload_excel()
        errors = check_func()
        # 清空结果显示区域
        self.result_area.config(state=tk.NORMAL)
        self.result_area.delete(1.0, tk.END)

        # 显示检查结果
        self.result_area.insert(tk.END, f"{check_name}:\n")
        if not errors:
            self.result_area.insert(tk.END, "  PASS\n", "green")
            if check_name == "序列检查":
                self.btn_sequence_revise.config(state=tk.DISABLED)
        else:
            self.result_area.insert(tk.END, "FAIL:\n", "red")
            self.result_area.insert(tk.END, "错误位置:\n", "red")
            for error in errors:
                self.result_area.insert(tk.END, f" 第 {error[0]} 行, 错误值: {error[2]}\n", "red")

            # 序列检查有问题则启用序号校正按钮
            if check_name == "序列检查":
                self.btn_sequence_revise.config(state=tk.NORMAL)
                self.sequence_error = check_sequence_for_fix(self.df)

        self.result_area.insert(tk.END, "\n")

        self.result_area.config(state=tk.DISABLED)

    def fix_sequence_errors(self, errors):
        """
        修正序列错误，并保存到原文件的合并单元格中
        :param errors: 错误列表，格式为 [(行号, 列名, 错误值), ...]
        :return:
        """
        try:
            fix_sequence(self.file_path, errors)
            messagebox.showinfo("success", "序列已校正并保存到modified_excel_file.xlsx！")
        except Exception as e:
            messagebox.showerror("fail!", f"校正序列失败{e}")


class AnalyseFrame(tk.Frame):
    """
    Analyse界面
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.create_widgets()
        self.file_path_analyse = None
        self.df1 = None
        self.canvas = None  # 用于保存画布对象
        self.ic_values = None
        self.module_values = None
        self.glass_values = None
        self.flash_values = None
        self.year_values = None
        # 用于模组厂分析分页
        self.current_page = 0
        self.total_pages = 1
        self.module_data = None
        self.module_title_conditions = None
        self.animation = None  # 用于存储动画对象
        self.fig = None  # 用于存储图形对象
        self.ax = None  # 用于存储坐标轴对象

    def create_widgets(self):
        """

        :return:
        """
        # 文件路径显示区域
        self.label1_file_path = tk.Label(self, text="文 件 路 径:")
        self.label1_file_path.place(x=10, y=10)

        self.text1_file_path = tk.Entry(self, state="readonly")
        self.text1_file_path.place(x=90, y=10, width=320 * multiple + 20, height=25)

        # 文件选择按钮
        self.btn_open_analyse = tk.Button(self, text="...", width=3, bg='Gainsboro', command=self.analyse_open_file)
        self.btn_open_analyse.place(x=420 * multiple, y=10)

        # 将数据分析及其下拉框框起来
        group_frame = tk.LabelFrame(self, bd=2, relief="groove", padx=5, pady=5)
        group_frame.place(x=460 * multiple, y=10, width=130, height=90)
        # 数据分析按键
        self.btn_analyse = tk.Button(self, text="数据分析", width=15, command=self.get_xl_to_analyse, state=tk.DISABLED)
        self.btn_analyse.place(x=465 * multiple, y=15)

        self.analyse_values = ["ALL", 'IC型号', '模组厂', '玻璃厂', 'Flash', '年度']
        self.analyse_dropdown = ttk.Combobox(self, values=self.analyse_values, state="readonly")
        self.analyse_dropdown.current(0)
        self.analyse_dropdown.place(x=475 * multiple, y=60, width=100)

        self.analyse_dropdown.bind("<<ComboboxSelected>>", self.choose_status)

        # 下拉框区域
        self.label_ic = tk.Label(self, text="IC型号:", anchor="e")
        self.label_ic.place(x=0, y=50, width=60)
        self.ic_values = ["ALL", '7272', '7202H', '7202M']
        self.ic_dropdown = ttk.Combobox(self, values=self.ic_values, state="readonly")
        self.ic_dropdown.current(0)
        self.ic_dropdown.place(x=60 * multiple, y=50, width=80)

        self.label_module = tk.Label(self, text="模组厂:", anchor="e")
        self.label_module.place(x=150 * multiple, y=50, width=60)
        self.module_values = ['ALL', '海菲', '信利', '易快来', '联创', '创维', '同兴达', '海盛捷', '壹星', '精卓', '三龙', '合力泰', '华显', '维立',
                              '亿华', '立德', '晶泰', '德智欣', '中光电', '沛宏', '欣欣光电', '众铭安', '天正达', '瑞恒光电', '清创高', '汉龙时代',
                              '晶胜通', '龙煜', '金宏光电', '威达光电', '惠科', '华视', '正金晶光电', '华映', '长信新显', '宏凯',
                              '泰启', '百业', '共赢', '德实', '京龙', '如新电子', '皓显', '大通显示', '高展', '康华显通', '煜鑫', '宏利超显', '菲触显视',
                              '轩达', '钜沣', '鹰芒技术', '德普特', '元格', '亿普拉斯']
        self.module_dropdown = ttk.Combobox(self, values=self.module_values, state="readonly")
        self.module_dropdown.current(0)
        self.module_dropdown.place(x=210 * multiple, y=50, width=80)

        self.label_glass = tk.Label(self, text="玻璃厂:", anchor="e")
        self.label_glass.place(x=300 * multiple, y=50, width=60)
        self.glass_values = ['ALL', 'CSOT', 'TM', 'TRULY', 'BOE', 'CTO', 'PANDA', 'SHARP', 'MDT', 'HKC', 'HSD', 'INX',
                             'IVO', 'CTC']
        self.glass_dropdown = ttk.Combobox(self, values=self.glass_values, state="readonly")
        self.glass_dropdown.current(0)
        self.glass_dropdown.place(x=360 * multiple, y=50, width=80)

        self.label_flash = tk.Label(self, text="Flash:", anchor="e")
        self.label_flash.place(x=0, y=80, width=60)
        self.flash_values = ['ALL', 'Y', 'N']
        self.flash_dropdown = ttk.Combobox(self, values=self.flash_values, state="readonly")
        self.flash_dropdown.current(0)
        self.flash_dropdown.place(x=60 * multiple, y=80, width=80)

        self.label_year = tk.Label(self, text="年度:", anchor="e")
        self.label_year.place(x=150 * multiple, y=80, width=60)
        self.year_values = ['ALL', '2022', '2023', '2024', '2025']
        self.year_dropdown = ttk.Combobox(self, values=self.year_values, state="readonly")
        self.year_dropdown.current(0)
        self.year_dropdown.place(x=210 * multiple, y=80, width=80)

        #模组分页设置
        # 初始化分页控件
        self.page_label = tk.Label(self)
        self.btn_prev = tk.Button(self, text="◀", width=2)
        self.btn_next = tk.Button(self, text="▶", width=2)

        # # 初始隐藏分页控件
        # self.page_label.place_forget()
        # self.btn_prev.place_forget()
        # self.btn_next.place_forget()

    def analyse_open_file(self):
        """

        :return:
        """
        # 打开文件选择对话框
        self.file_path_analyse = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])

        if self.file_path_analyse:
            # 更新文件路径
            self.text1_file_path.config(state="normal")
            self.text1_file_path.delete(0, tk.END)
            self.text1_file_path.insert(0, self.file_path_analyse)
            self.text1_file_path.config(state="readonly")

            # 启用数据分析键
            self.btn_analyse.config(state=tk.NORMAL)

            # 重新选择文件则清除画布
            self.clear_plot()

    def analyse_reload_excel(self):
        """
        加载文件
        :return:
        """
        try:
            self.df1 = read_excel_for_analyse(self.file_path_analyse, sheet_name='MP Project List（internal）')
        except Exception as e:
            self.df1 = None
            messagebox.showerror("错误", f"读取Excel文件失败: {e}")

    def choose_status(self, event=None):
        choose = self.analyse_dropdown.get()
        if choose == 'ALL':
            self.ic_dropdown.config(state=tk.NORMAL)
            self.module_dropdown.config(state=tk.NORMAL)
            self.glass_dropdown.config(state=tk.NORMAL)
            self.flash_dropdown.config(state=tk.NORMAL)
            self.year_dropdown.config(state=tk.NORMAL)

        elif choose == 'IC型号':
            self.ic_dropdown.config(state=tk.DISABLED)
            self.module_dropdown.config(state=tk.NORMAL)
            self.glass_dropdown.config(state=tk.NORMAL)
            self.flash_dropdown.config(state=tk.NORMAL)
            self.year_dropdown.config(state=tk.NORMAL)

        elif choose == '模组厂':
            self.ic_dropdown.config(state=tk.NORMAL)
            self.module_dropdown.config(state=tk.DISABLED)
            self.glass_dropdown.config(state=tk.NORMAL)
            self.flash_dropdown.config(state=tk.NORMAL)
            self.year_dropdown.config(state=tk.NORMAL)

        elif choose == '玻璃厂':
            self.ic_dropdown.config(state=tk.NORMAL)
            self.module_dropdown.config(state=tk.NORMAL)
            self.glass_dropdown.config(state=tk.DISABLED)
            self.flash_dropdown.config(state=tk.NORMAL)
            self.year_dropdown.config(state=tk.NORMAL)

        elif choose == 'Flash':
            self.ic_dropdown.config(state=tk.NORMAL)
            self.module_dropdown.config(state=tk.NORMAL)
            self.glass_dropdown.config(state=tk.NORMAL)
            self.flash_dropdown.config(state=tk.DISABLED)
            self.year_dropdown.config(state=tk.NORMAL)
        elif choose == '年度':
            self.ic_dropdown.config(state=tk.NORMAL)
            self.module_dropdown.config(state=tk.NORMAL)
            self.glass_dropdown.config(state=tk.NORMAL)
            self.flash_dropdown.config(state=tk.NORMAL)
            self.year_dropdown.config(state=tk.DISABLED)

    def get_xl_to_analyse(self):
        """

        :return:
        """
        try:
            self.analyse_reload_excel()  # 每次点击数据分析可重新加载文件
            choose = self.analyse_dropdown.get()
            self.clear_plot()  # 每次选择先清除上一次直方图

            # 重置分页状态
            self.current_page = 0
            self.total_pages = 1
            # 其他分析类型不需要分页
            self.update_page_buttons()  # 这将隐藏分页按钮

            if choose == 'ALL':
                ic_type = self.ic_dropdown.get()
                factory = self.module_dropdown.get()
                glass = self.glass_dropdown.get()
                flash = self.flash_dropdown.get()
                year = self.year_dropdown.get()

                conditions, dist_stats = statistic_project_status(self.df1, ic_type, factory, glass, flash, year)

                # 检查DataFrame是否为空
                if dist_stats.empty or len(dist_stats) == 0:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")
                    return
                if len(dist_stats.columns) >= 2:
                    project_counts = dist_stats.iloc[:, 1]  # 获取第二列数据
                else:
                    project_counts = dist_stats.iloc[:, 0]  # 如果只有一列，使用第一列  传入的空Series

                if len(project_counts) != 0:
                    fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                    f_plot = fig.add_subplot(111)  # 划分区域
                    self.canvas = FigureCanvasTkAgg(fig, master=self)
                    self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                    f_plot.clear()  # 刷新
                    ax = sns.barplot(
                        x='Update_Count',
                        y='Project_Count',
                        data=dist_stats,
                        linewidth=1
                    )
                    # 添加数据标签
                    for p in ax.patches:
                        ax.annotate(
                            f"{int(p.get_height())}",
                            (p.get_x() + p.get_width() / 2., p.get_height()),
                            ha='center', va='center',
                            xytext=(0, 5),
                            textcoords='offset points',
                            fontsize=10,
                            weight='bold'
                        )
                    title = "项目更新次数分布统计\n(" + " | ".join(conditions) + ")" if conditions else "项目更新次数分布统计（ALL）"
                    plt.title(title, fontsize=14, pad=10)
                    plt.xlabel('更新次数', fontsize=12)
                    plt.ylabel('项目数量', fontsize=12)
                    plt.xticks(fontsize=10)
                    plt.yticks(fontsize=10)

                    # 自动调整y轴范围，留出标签空间
                    max_value = max((dist_stats.iloc[:, 1]))
                    plt.ylim(0, max_value * 1.15)
                else:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")


            elif choose == 'IC型号':

                factory = self.module_dropdown.get()
                glass = self.glass_dropdown.get()
                flash = self.flash_dropdown.get()
                year = self.year_dropdown.get()

                title_conditions, ic_counts = statistic_ic_projects(self.df1, factory, glass, flash, year)
                # 绘制图表
                if int(ic_counts.sum()) != 0:
                    fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                    f_plot = fig.add_subplot(111)  # 划分区域
                    self.canvas = FigureCanvasTkAgg(fig, master=self)
                    self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                    f_plot.clear()  # 刷新

                    bars = plt.bar(ic_counts.index, ic_counts.values, width=0.6)
                    # 添加数据标签
                    for bar in bars:
                        height = bar.get_height()
                        plt.text(bar.get_x() + bar.get_width() / 2., height,
                                 f'{int(height)}',
                                 ha='center', va='bottom', fontsize=11, weight='bold')
                    title = "IC型号项目数量统计"
                    if title_conditions:
                        title += "\n(" + " | ".join(title_conditions) + ")"

                    plt.title(title, fontsize=14, pad=10)
                    plt.xlabel('IC型号', fontsize=12)
                    plt.ylabel('项目数量', fontsize=12)
                    plt.xticks(fontsize=10)
                    plt.yticks(fontsize=10)

                    # 自动调整y轴范围，留出标签空间
                    max_value = max(ic_counts.values)
                    plt.ylim(0, max_value * 1.15)
                else:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")

            elif choose == '模组厂':
                ic_type = self.ic_dropdown.get()
                glass = self.glass_dropdown.get()
                flash = self.flash_dropdown.get()
                year = self.year_dropdown.get()
                title_conditions, factory_counts = statistic_module_factory(self.df1, ic_type, glass, flash, year)
                self.module_data = factory_counts  # 保存模组厂数据

                if int(factory_counts.sum()) > 0:
                    # 计算总页数
                    self.total_pages = (len(factory_counts) + 19) // 20
                    if self.total_pages > 1:
                        self.show_module_page(title_conditions=title_conditions)
                    elif self.total_pages == 1:
                        # self.clear_plot()
                        # 绘制图表
                        fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                        f_plot = fig.add_subplot(111)  # 划分区域
                        self.canvas = FigureCanvasTkAgg(fig, master=self)
                        self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                        f_plot.clear()  # 刷新
                        bars = plt.barh(factory_counts.index, factory_counts.values)

                        max_value = max(factory_counts.values)
                        if max_value > 3:
                            # 添加数据标签
                            for bar in bars:
                                width = bar.get_width()
                                plt.text(width + 0.2,  # x位置：宽度值+0.2的偏移
                                         bar.get_y() + bar.get_height() / 2,  # y位置：居中
                                         f'{int(width)}',
                                         va='center',  # 垂直居中
                                         ha='left',  # 水平左对齐
                                         fontsize=10)
                        else:
                            # 添加数据标签
                            for bar in bars:
                                width = bar.get_width()
                                plt.text(width,  # x位置不偏移
                                         bar.get_y() + bar.get_height() / 2,  # y位置：居中
                                         f'{int(width)}',
                                         va='center',  # 垂直居中
                                         ha='left',  # 水平左对齐
                                         fontsize=10)

                        title = "模组厂项目数量统计"
                        if title_conditions:
                            title += "\n(" + " | ".join(title_conditions) + ")"

                        plt.title(title, fontsize=14, pad=10)
                        plt.ylabel('模组厂', fontsize=12)
                        plt.xlabel('项目数量', fontsize=12)
                        # plt.xticks(rotation=90, fontsize=10)
                        plt.xticks(fontsize=10)
                        plt.yticks(fontsize=10)

                        # 自动调整x轴范围，留出标签空间
                        # max_value = max(factory_counts.values)
                        plt.xlim(0, max_value * 1.15)
                elif int(factory_counts.sum()) == 0:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")

            elif choose == '玻璃厂':
                ic_type = self.ic_dropdown.get()
                factory = self.module_dropdown.get()
                flash = self.flash_dropdown.get()
                year = self.year_dropdown.get()

                title_conditions, glass_counts = statistic_glass_projects(self.df1, ic_type, factory, flash, year)

                if int(glass_counts.sum()) != 0:
                    # 绘制图表
                    fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                    f_plot = fig.add_subplot(111)  # 划分区域
                    self.canvas = FigureCanvasTkAgg(fig, master=self)
                    self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                    f_plot.clear()  # 刷新

                    bars = plt.bar(glass_counts.index, glass_counts.values)

                    # 添加数据标签
                    for bar in bars:
                        height = bar.get_height()
                        plt.text(bar.get_x() + bar.get_width() / 2., height,
                                 f'{int(height)}',
                                 ha='center', va='bottom', fontsize=10)

                    title = "玻璃类型项目数量统计"
                    if title_conditions:
                        title += "\n(" + " | ".join(title_conditions) + ")"

                    plt.title(title, fontsize=14, pad=10)
                    plt.xlabel('玻璃类型', fontsize=12)
                    plt.ylabel('项目数量', fontsize=12)
                    plt.xticks(rotation=45, fontsize=10)
                    plt.yticks(fontsize=10)

                    # 自动调整y轴范围，留出标签空间
                    max_value = max(glass_counts.values)
                    plt.ylim(0, max_value * 1.15)
                else:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")

            elif choose == 'Flash':
                ic_type = self.ic_dropdown.get()
                factory = self.module_dropdown.get()
                glass = self.glass_dropdown.get()
                year = self.year_dropdown.get()
                title_conditions, flash_counts = statistic_flash_projects(self.df1, ic_type, factory, glass, year)

                if int(flash_counts.sum()) != 0:
                    # 绘制图表
                    fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                    f_plot = fig.add_subplot(111)  # 划分区域
                    self.canvas = FigureCanvasTkAgg(fig, master=self)
                    self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                    f_plot.clear()  # 刷新

                    bars = plt.bar(flash_counts.index, flash_counts.values, width=0.6)

                    # 添加数据标签
                    for bar in bars:
                        height = bar.get_height()
                        plt.text(bar.get_x() + bar.get_width() / 2., height,
                                 f'{int(height)}',
                                 ha='center', va='bottom', fontsize=12, weight='bold')

                    title = "带Flash vs 不带Flash项目数量统计"
                    if title_conditions:
                        title += "\n(" + " | ".join(title_conditions) + ")"

                    plt.title(title, fontsize=14, pad=10)
                    plt.xlabel('是否带Flash', fontsize=12)
                    plt.ylabel('项目数量', fontsize=12)
                    plt.xticks(fontsize=10)
                    plt.yticks(fontsize=10)

                    # 自动调整y轴范围，留出标签空间
                    max_value = max(flash_counts.values)
                    plt.ylim(0, max_value * 1.15)
                else:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")

            elif choose == '年度':
                ic_type = self.ic_dropdown.get()
                factory = self.module_dropdown.get()
                glass = self.glass_dropdown.get()
                flash = self.flash_dropdown.get()

                title_conditions, year_counts = statistic_project_by_year(self.df1, ic_type, factory, glass, flash)

                if int(year_counts.sum()) != 0:
                    # 绘制图表
                    fig = plt.figure(figsize=(7.2, 5.5), dpi=100)
                    f_plot = fig.add_subplot(111)  # 划分区域
                    self.canvas = FigureCanvasTkAgg(fig, master=self)
                    self.canvas.get_tk_widget().place(x=0, y=120)  # 放置位置
                    f_plot.clear()  # 刷新

                    bars = plt.bar(year_counts.index, year_counts.values, width=0.6)

                    # 添加数据标签
                    for bar in bars:
                        height = bar.get_height()
                        plt.text(bar.get_x() + bar.get_width() / 2., height,
                                 f'{int(height)}',
                                 ha='center', va='bottom', fontsize=12, weight='bold')

                    title = "年度项目数量统计"
                    if title_conditions:
                        title += "\n(" + " | ".join(title_conditions) + ")"

                    plt.title(title, fontsize=14, pad=10)
                    plt.xlabel('年度', fontsize=12)
                    plt.ylabel('项目数量', fontsize=12)
                    plt.xticks(fontsize=10)
                    plt.yticks(fontsize=10)

                    # 自动调整y轴范围，留出标签空间
                    max_value = max(year_counts.values)
                    plt.ylim(0, max_value * 1.15)
                else:
                    messagebox.showinfo("提示", "该筛选条件下项目数量为0")

            # # 设置网格线
            # f_plot.grid(True, linestyle='--', alpha=0.6)
            # 使直方图填充整个画布
            plt.tight_layout()
            if self.canvas:
                self.canvas.draw()

        except Exception as e:
            messagebox.showerror("错误", f"数据分析时发生错误: {str(e)}")
            self.clear_plot()
            self.total_pages = 1
            self.update_page_buttons()

    def show_module_page(self, title_conditions, direction=0):
        """
        显示当前页码的模组厂数据
        :param direction: 0-无动画, 1-向右滑动(下一页), -1-向左滑动(上一页)
        :return:
        """
        if self.module_data is None:
            return

        start_idx = self.current_page * 20
        end_idx = min((self.current_page + 1) * 20, len(self.module_data))
        current_data = self.module_data[start_idx:end_idx]

        # 创建图形和坐标轴(如果不存在)
        if self.fig is None:
            self.fig = plt.figure(figsize=(7.2, 5.1), dpi=100)
            self.ax = self.fig.add_subplot(111)

        # 清除之前的画布
        if self.canvas:
            self.clear_plot()

        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().place(x=0, y=120)
        self.ax.clear()

        fixed_height = 0.6  # 固定高度值
        # 计算间距 - 确保最后一页条目少时也能保持相同高度
        n_bars = len(current_data)
        if n_bars < 20:
            # 调整y轴范围使条形高度保持一致
            self.ax.set_ylim(-0.5, 19.5)  # 保持20个条形的空间

        # 绘制柱状图
        bars = self.ax.barh(
            y=range(n_bars),
            width=current_data.values,
            height=fixed_height)

        # 设置y轴刻度标签为实际模组厂名称
        self.ax.set_yticks(range(n_bars))
        self.ax.set_yticklabels(current_data.index)

        # 添加数据标签
        for bar in bars:
            width = bar.get_width()
            self.ax.text(width + 0.2,
                         bar.get_y() + bar.get_height() / 2,
                         f'{int(width)}',
                         va='center',
                         ha='left',
                         fontsize=10)

        title = f"模组厂项目数量统计"
        self.module_title_conditions = title_conditions
        if self.module_title_conditions:
            title += "\n(" + " | ".join(self.module_title_conditions) + ")"
            # 调整子图位置和标题位置，为多行标题留出空间
            self.fig.subplots_adjust(top=0.86)  # 增加顶部空间
        else:
            self.fig.subplots_adjust(top=0.9)  # 默认顶部空间

        self.ax.set_title(title, fontsize=14, pad=10)
        self.ax.set_ylabel('模组厂', fontsize=12)
        self.ax.set_xlabel('项目数量', fontsize=12)

        # 自动调整x轴范围
        global_max = max(self.module_data.values) * 1.15
        self.ax.set_xlim(0, global_max)

        # 固定x轴刻度
        max_ticks = 10
        interval = max(1, int(global_max / max_ticks))
        self.ax.set_xticks(range(0, int(global_max) + 1, interval))

        self.ax.tick_params(axis='x', labelsize=10)
        self.ax.tick_params(axis='y', labelsize=10)

        if self.canvas:
            # 如果有动画需求
            if direction != 0:
                self.create_slide_animation(direction)
            else:
                self.canvas.draw()

        # 更新翻页按钮状态
        self.update_page_buttons()

    def update_page_buttons(self):
        """
        更新翻页按钮状态
        :return:
        """
        # 隐藏所有分页控件
        self.page_label.place_forget()
        self.btn_prev.place_forget()
        self.btn_next.place_forget()

        # 只有多于1页且是模组厂分析时才显示分页按钮
        if self.total_pages > 1 and self.analyse_dropdown.get() == '模组厂':
            # 更新文本和状态
            self.page_label.config(text=f"第{self.current_page + 1}页/共{self.total_pages}页")
            self.btn_prev.config(
                command=self.prev_page,
                state=tk.NORMAL if self.current_page > 0 else tk.DISABLED
            )
            self.btn_next.config(
                command=self.next_page,
                state=tk.NORMAL if self.current_page < self.total_pages - 1 else tk.DISABLED
            )

            # 放置控件
            self.page_label.place(x=460 * multiple, y=525 * multiple)
            self.btn_prev.place(x=530 * multiple, y=525 * multiple)
            self.btn_next.place(x=550 * multiple, y=525 * multiple)

    def create_slide_animation(self, direction):
        """创建滑动动画
        direction: 1-向右滑动(下一页), -1-向左滑动(上一页)
        """
        # 停止之前的动画
        if self.animation:
            self.animation.event_source.stop()

        # 初始位置设置
        initial_offset = 0.1 * direction  # 初始偏移量(10%的宽度)
        self.ax.set_position([0.5 + initial_offset, 0.15, 0.85, 0.8])  # [left, bottom, width, height]

        # 动画帧数
        frames = 10
        step = -initial_offset / frames

        def update(frame):
            offset = initial_offset + step * frame
            self.ax.set_position([0.5 + offset, 0.15, 0.85, 0.8])
            return self.ax,

        # 创建动画
        self.animation = animation.FuncAnimation(
            self.fig,
            update,
            frames=frames,
            interval=20,  # 每帧间隔20ms
            blit=True
        )

        self.canvas.draw()

    def prev_page(self):
        """
        上一页
        """
        if self.current_page > 0:
            self.current_page -= 1
            self.show_module_page(self.module_title_conditions)

    def next_page(self):
        """
        下一页
        """
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.show_module_page(self.module_title_conditions)

    def clear_plot(self):
        """
        清除直方图
        :return:
        """
        if self.canvas:
            self.canvas.get_tk_widget().destroy()  # 销毁画布
            self.canvas = None



class MainApp:
    """
    主程序界面，用于切换create和 Analyse界面
    """

    def __init__(self, root):
        self.root = root
        self.root.title("MP List Check Tool_v4.1_20250409")
        # self.root.geometry("600x580")  # 设置窗口大小
        self.root.resizable(0, 0)

        # 创建标签页容器
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        # 创建Check标签页
        self.check_frame = CheckFrame(self.notebook)
        self.notebook.add(self.check_frame, text='   Check   ')

        # 创建Analyse标签页
        self.analyse_frame = AnalyseFrame(self.notebook)
        self.notebook.add(self.analyse_frame, text='   Analyse   ')

        # 绑定标签页切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # 配置标签样式
        self.style = ttk.Style()
        self.style.configure('TNotebook', background='#f0f0f0')
        self.style.configure('TNotebook.Tab',
                             font=('Arial', 10, 'bold'),
                             padding=(20, 5),
                             background='#e0e0e0',
                             foreground='#444')
        self.style.map('TNotebook.Tab',
                       background=[('selected', '#ffffff')],
                       expand=[('selected', [1, 1, 1, 0])])

    def on_tab_changed(self, event):
        """
        标签页切换事件处理
        :return:
        """
        selected_tab = self.notebook.select()
        tab_text = self.notebook.tab(selected_tab, "text")

        if tab_text.strip() == "Analyse":
            # 切换到Analyse界面时，确保直方图显示
            if self.analyse_frame.canvas:
                self.analyse_frame.canvas.get_tk_widget().place(x=0, y=120)
            # pass
        elif tab_text.strip() == "Check":
            # 切换到其他界面时，清除直方图
            # self.analyse_frame.clear_plot()
            pass


if __name__ == "__main__":
    """
    运行函数
    """
    root = tk.Tk()
    # 计算屏幕中央的位置
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    width = 600 * multiple
    heigh = 580 * multiple
    # 将窗口居中显示
    root.geometry('%dx%d+%d+%d' % (width, heigh, (screenwidth - width) / 2, (screenheight - heigh) / 2))
    app = MainApp(root)
    root.mainloop()
