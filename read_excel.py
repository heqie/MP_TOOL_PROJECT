#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT 
@File    ：read_excel.py
@Author  ：zxh
@Date    ：2025/3/19 16:21 
'''
import pandas as pd
from tkinter import messagebox


def read_excel1(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df


def read_excel_with_merged_cells(file_path, sheet_name):
    """
    完成MP list文件读入并处理合并单元格
    :file_path: 文件路径
    :sheet_name: 文件页
    :return: df 处理后的DataFrame
    """
    try:
        with pd.ExcelFile(file_path) as excel_file:
            all_sheet_names = excel_file.sheet_names

            # 如果未指定 Sheet 名称，使用第一个 Sheet
            if sheet_name is None:
                sheet_name = all_sheet_names[0]

            # 检查指定的 Sheet 是否存在
            if sheet_name not in all_sheet_names:
                raise ValueError(f"Sheet '{sheet_name}' 不存在")

            # 读取指定的 Sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            if sheet_name == 'MP Project List（internal）':
                df = pd.DataFrame(df)
                # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)
                # print(df.iloc[:, 0])
                df.iloc[:, 2] = df.iloc[:, 2].ffill()
                # 替换项目号中的特定值
                cleaned_project_numbers = df.iloc[:, 2]

                # 读取对比文件
                valid_data_file = 'compare_information.xlsx'
                try:
                    valid_data = read_compare_excel(valid_data_file, sheet_names=None)
                    if valid_data is not None and not valid_data.empty:
                        change_before_project = valid_data.iloc[:, 4].dropna().tolist()
                        change_after_project = valid_data.iloc[:, 5].dropna().tolist()
                    else:
                        change_before_project = []
                        change_after_project = []
                        messagebox.showwarning("警告", "对比文件为空或无效")
                except Exception as e:
                    change_before_project = []
                    change_after_project = []
                    messagebox.showerror("错误", f"读取对比文件失败: {e}")

                change_mapping = dict(zip(change_before_project, change_after_project))
                for idx, project_number in enumerate(cleaned_project_numbers):
                    if project_number in change_mapping:
                        cleaned_project_numbers[idx] = change_mapping[project_number]
                        # print(f"项目号 {project_number} 已替换为 {cleaned_project_numbers[idx]}")

                # 合并的单元格为空则左上角填入0
                replace_num = int(0)
                # 序号
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 0].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1]:
                        df.iloc[idx, 0] = replace_num

                # ic
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 6].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1]:
                        df.iloc[idx, 6] = replace_num

                # 玻璃
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 7].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1]:
                        df.iloc[idx, 7] = replace_num

                # 接口
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 8].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1]:
                        df.iloc[idx, 8] = replace_num

                # flash
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 9].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1]:
                        df.iloc[idx, 9] = replace_num

                # 向下填充NAN值
                df.iloc[:, 0] = df.iloc[:, 0].ffill()
                df.iloc[:, 2] = df.iloc[:, 2].ffill()
                df.iloc[:, 6] = df.iloc[:, 6].ffill()
                df.iloc[:, 7] = df.iloc[:, 7].ffill()
                df.iloc[:, 8] = df.iloc[:, 8].ffill()
                df.iloc[:, 9] = df.iloc[:, 9].ffill()
                df.iloc[:, 10] = df.iloc[:, 10].ffill()

                # 将空字符串（excel单元格为文本类型）单元格数据替换为0
                df = df.replace('', '0')
                # # 填充非文本类型单元格数据为0
                df = df.fillna(0)

        return df
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None


def read_compare_excel(file_path, sheet_names=None):
    """
    对于compare文件的处理
    :param file_path:文件路径
    :param sheet_names:文件页
    :return:df
    """
    try:
        # 使用 pandas.ExcelFile 获取所有工作表的名称
        with pd.ExcelFile(file_path) as excel_file:
            all_sheet_names = excel_file.sheet_names

            # 如果未指定 Sheet 名称，使用第一个 Sheet
            if sheet_names is None:
                sheet_names = all_sheet_names[0]

            # 检查指定的 Sheet 是否存在
            if sheet_names not in all_sheet_names:
                raise ValueError(f"Sheet '{sheet_names}' 不存在")

            # 读取指定的 Sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_names)

        # 将空字符串（excel单元格为文本类型）单元格数据替换为0
        df = df.replace('', '0')
        # 删除字符串中的空格和换行符
        df = df.replace({'\n': '', ' ': ''}, regex=True)

        return df

    except Exception as e:
        print(f"读取compare Excel文件失败: {e}")
        return None


def read_excel_for_analyse(file_path, sheet_name):
    """
    读取并处理excel用于更新次数统计
    :param file_path:文件路径
    :param sheet_name:文件页
    :return: data
    """
    try:
        with pd.ExcelFile(file_path) as excel_file:
            all_sheet_names = excel_file.sheet_names

            # 如果未指定 Sheet 名称，使用第一个 Sheet
            if sheet_name is None:
                sheet_name = all_sheet_names[0]

            # 检查指定的 Sheet 是否存在
            if sheet_name not in all_sheet_names:
                raise ValueError(f"Sheet '{sheet_name}' 不存在")

            # 读取指定的 Sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            if sheet_name == 'MP Project List（internal）':
                df = pd.DataFrame(df)
                # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)

                df.iloc[:, 2] = df.iloc[:, 2].ffill()
                # 模组厂
                df.iloc[:, 5] = df.iloc[:, 5].ffill()
                df.replace({
                    '汉龙': '汉龙时代',
                    '轩达百业': '轩达',
                    '轩达(百业代工）': '轩达',
                    '轩达（星星科技代工）': '轩达',
                    '惠科(华锐代)': '惠科',
                    '重庆联创': '联创',
                    '两江联创': '联创',
                    '万年联创': '联创',
                    '大通': '大通显示',
                    '菲触': '菲触显视',
                    '宏利': '宏利超显',
                    '瑞恒': '瑞恒光电'
                }, inplace=True)

                df.iloc[:, 10] = df.iloc[:, 10].ffill()

                # ic
                df.iloc[:, 6] = df.iloc[:, 6].ffill()
                # 玻璃
                df.iloc[:, 7] = df.iloc[:, 7].ffill()

                # flash
                df.iloc[:, 9] = df.iloc[:, 9].ffill()

                # 将空字符串（excel单元格为文本类型）单元格数据替换为0
                df = df.replace('', '0')
                # # 填充非文本类型单元格数据为0
                df = df.fillna(0)
                # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)

        return df
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None
