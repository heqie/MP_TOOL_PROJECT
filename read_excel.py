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
import re


def remove_parentheses(text):
    """
     删除字符串中的括号及其里面的字符
    :param text:
    :return:
    """
    # 英文括号模式
    english_pattern = r"\([^()]*\)"
    # 中文括号模式
    chinese_pattern = r"（[^（）]*）"

    # 先处理中文括号
    text = re.sub(chinese_pattern, "", text)
    # 再处理英文括号
    text = re.sub(english_pattern, "", text)
    return text


def remove_after_underscore(text):
    """
    删除字符串中的"_"及其之后的字符
    :param text:
    :return:
    """
    # 匹配下划线及其后的所有字符
    pattern = r"_.*"
    result = re.sub(pattern, "", text)
    return result


def convert_to_uppercase(text):
    """
    将字符串全部转化为大写
    :param text:
    :return:
    """
    return text.upper()


def remove_english_keep_chinese(text):
    """
    删除字符串中的英文（包括字母、数字、英文标点），只保留中文
    :param text: 输入字符串
    :return: 只包含中文的字符串
    """
    # 匹配非中文字符（包括英文、数字、英文标点、空格等）
    pattern = r'[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]'
    text = re.sub(pattern, '', text)
    return text


def remove_after_digits_keep_separators(text):
    """
    去除字符串中数字后面的所有字符，保留数字以及由-或.连接的数字组合
    :param text: 输入字符串
    :return: 处理后的字符串
    """
    if not isinstance(text, str):
        text = str(text)
    # 先将"华星"替换为"CSOT"
    text = text.replace("华星", "CSOT")

    # 匹配数字以及由-或.连接的数字组合
    pattern = r'(\d+(?:[.-]\d+)*)\s*.*$'

    result = re.sub(pattern, r'\1', text)

    # 如果结果和原字符串相同，说明没有匹配到数字
    if result == text:
        # 检查是否包含数字
        if not re.search(r'\d', text):
            return text

    return result


def remove_characters_around_hyphen_regex(text):
    """
    使用正则表达式删除'-'前后的一个字符以及'-'本身
    :param text: 输入字符串
    :return: 处理后的字符串
    """
    import re

    if not isinstance(text, str):
        text = str(text)

    # 首先处理连字符在中间的情况（前后都有字符）
    pattern1 = r'.?-.?'

    # 使用循环处理，因为一次替换后可能产生新的连字符组合
    old_text = text
    while True:
        new_text = re.sub(pattern1, '', old_text)
        if new_text == old_text:
            break
        old_text = new_text

    return old_text


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
                #
                # # 读取对比文件
                # valid_data_file = 'compare_information.xlsx'
                # try:
                #     valid_data = read_compare_excel(valid_data_file, sheet_names='Sheet1')
                #     if valid_data is not None and not valid_data.empty:
                #         change_before_project = valid_data.iloc[:, 4].dropna().tolist()
                #         change_after_project = valid_data.iloc[:, 5].dropna().tolist()
                #     else:
                #         change_before_project = []
                #         change_after_project = []
                #         messagebox.showwarning("警告", "对比文件为空或无效")
                # except Exception as e:
                #     change_before_project = []
                #     change_after_project = []
                #     messagebox.showerror("错误", f"读取对比文件失败: {e}")
                #
                # change_mapping = dict(zip(change_before_project, change_after_project))
                # for idx, project_number in enumerate(cleaned_project_numbers):
                #     if project_number in change_mapping:
                #         cleaned_project_numbers[idx] = change_mapping[project_number]
                #         # print(f"项目号 {project_number} 已替换为 {cleaned_project_numbers[idx]}")
                # cleaned_project_numbers = cleaned_project_numbers
                # 合并的单元格为空则左上角填入0
                replace_num = int(0)
                cleaned_squence = df.iloc[:, 0].fillna('')
                # 序号
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 0].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 0] = replace_num

                # 终端
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 3].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 3] = replace_num

                # ic
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 6].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 6] = replace_num

                # 玻璃
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 7].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 7] = replace_num

                # 接口
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 8].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 8] = replace_num

                # flash
                for idx, (project_number, sequence) in enumerate(
                        zip(cleaned_project_numbers, df.iloc[:, 9].fillna(''))):
                    if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[
                        idx - 1] and sequence != cleaned_squence[idx - 1]:
                        df.iloc[idx, 9] = replace_num

                # 向下填充NAN值
                df.iloc[:, 0] = df.iloc[:, 0].ffill()
                df.iloc[:, 2] = df.iloc[:, 2].ffill()
                df.iloc[:, 3] = df.iloc[:, 3].ffill()
                df.iloc[:, 5] = df.iloc[:, 5].ffill()
                df.replace({
                    '汉龙': '汉龙时代',
                    '汉龙时代光电': '汉龙时代',
                    '轩达百业': '轩达',
                    '轩达(百业代工）': '轩达',
                    '轩达（星星科技代工）': '轩达',
                    '惠科(华锐代)': '惠科',
                    '重庆联创': '联创',
                    '两江联创': '联创',
                    '万年联创': '联创',
                    '万联': '联创',
                    '重联': '联创',
                    '大通': '大通显示',
                    '菲触': '菲触显视',
                    '宏利': '宏利超显',
                    '瑞恒': '瑞恒光电',
                    '湖南金宏光电': '金宏光电',
                    '新显': '长信新显',
                    '天山': '天山电子',
                }, inplace=True)
                df.iloc[:, 6] = df.iloc[:, 6].ffill()
                df.iloc[:, 7] = df.iloc[:, 7].ffill()
                df.iloc[:, 8] = df.iloc[:, 8].ffill()
                df.iloc[:, 9] = df.iloc[:, 9].ffill()
                df.iloc[:, 10] = df.iloc[:, 10].ffill()

                # 将空字符串（excel单元格为文本类型）单元格数据替换为0
                df = df.replace('', '0')
                # # 填充非文本类型单元格数据为0
                df = df.fillna(0)

                # debug处理后的文件
                # save_path = "processed_data.xlsx"
                # df.to_excel(save_path, index=False, sheet_name=sheet_name)

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
            # print(df.iloc[:, 7])
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

                # df.iloc[:, 1].astype(str).str.upper()
                df.iloc[:, 2] = df.iloc[:, 2].ffill()
                # # 替换项目号中的特定值
                # replace_num = int(0)
                # cleaned_squence = df.iloc[:, 0].fillna('')
                # cleaned_project_numbers = df.iloc[:, 2]
                # for idx, (project_number, sequence) in enumerate(
                #         zip(cleaned_project_numbers, df.iloc[:, 3].fillna(''))):
                #     if (pd.isna(sequence) or sequence == '') and project_number != cleaned_project_numbers[idx - 1] and sequence != cleaned_squence[idx - 1]:
                #         df.iloc[idx, 3] = replace_num
                df.iloc[:, 0] = df.iloc[:, 0].ffill()
                # 终端
                df.iloc[:, 3] = df.iloc[:, 3].ffill()
                # 模组厂
                df.iloc[:, 5] = df.iloc[:, 5].ffill()
                df.replace({
                    '汉龙': '汉龙时代',
                    '汉龙时代光电': '汉龙时代',
                    '轩达百业': '轩达',
                    '轩达(百业代工）': '轩达',
                    '轩达（星星科技代工）': '轩达',
                    '惠科(华锐代)': '惠科',
                    '重庆联创': '联创',
                    '两江联创': '联创',
                    '万年联创': '联创',
                    '万联': '联创',
                    '重联': '联创',
                    '大通': '大通显示',
                    '菲触': '菲触显视',
                    '宏利': '宏利超显',
                    '瑞恒': '瑞恒光电',
                    '湖南金宏光电': '金宏光电',
                    '新显': '长信新显',
                    '天山': '天山电子',
                }, inplace=True)

                df.iloc[:, 10] = df.iloc[:, 10].ffill()

                # ic
                df.iloc[:, 6] = df.iloc[:, 6].ffill()
                # 玻璃
                df.iloc[:, 7] = df.iloc[:, 7].ffill()

                # flash
                df.iloc[:, 9] = df.iloc[:, 9].ffill()

                # df.iloc[:, 1] = df.iloc[:, 1].replace('NAN', 'a')

                # 将空字符串（excel单元格为文本类型）单元格数据替换为0
                # df = df.replace('', '0')
                # # 填充非文本类型单元格数据为0
                for idx, value in enumerate(df.iloc[:, 1]):
                    if pd.isna(value):
                        df.iloc[idx, 1] = 'NULL'

                df = df.fillna(0)
                # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)

                # df.to_excel('output.xlsx', index=False)

        return df
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None


def read_excel_for_analyseKS(file_path, sheet_name):
    """
    读取并处理excel用于分析客诉记录统计
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
            if sheet_name == '客诉记录':

                df = pd.DataFrame(df)
                # # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)
                # 将合并单元格填充相同的内容
                for i in range(0, 14):
                    df.iloc[:, i] = df.iloc[:, i].ffill()

                # df.iloc[:, 12] = df.iloc[:, 12].astype(str).str.replace(r'\(.*?\)', '', regex=True)  # 英文括号
                # df.iloc[:, 12].replace(r"\(.*?\）", "", regex=True)  # 中文括号

                # 客诉类型去掉括号及其括号内容
                df.iloc[:, 12] = df.iloc[:, 12].apply(lambda x: remove_parentheses(str(x)))
                # 替代部分不规范字符
                df.replace({
                    'Y': '是',
                    'N': '否',
                    '大通': '大通显示',
                    '汉龙': '汉龙时代',
                    '立德/新显': '长信新显',
                    '新显': '长信新显',
                    'IC来料、玻璃': '玻璃',
                    '万联': '联创',
                    '重联': '联创',
                    # '误判(Tool+MP Bin)' :'误判',
                    # '误判(Tool)':'误判',
                    # '误判(MP Bin)':'误判'
                }, inplace=True)

                # 删除字符串中的空格和换行符
                df = df.replace({'\n': '', ' ': ''}, regex=True)

        return df
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None


def read_excel_for_analyseHN(file_path, sheet_name = None):
    """
    读取并处理excel用于华南项目统计
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
            df_2025 = pd.read_excel(excel_file, sheet_name='2025年TDDI',skiprows=[0])
            df_2025 = pd.DataFrame(df_2025)
            # # 删除字符串中的空格和换行符
            df_2025 = df_2025.replace({'\n': '', ' ': ''}, regex=True)

            # 去除模组厂字符串
            df_2025.iloc[:, 3] = df_2025.iloc[:, 3].apply(lambda x: remove_english_keep_chinese(str(x)))
            df_2025.iloc[:, 3] = df_2025.iloc[:, 3].apply(lambda x: remove_parentheses(str(x)))

            # 去除玻璃后面型号备注
            df_2025.iloc[:, 5] = df_2025.iloc[:, 5].apply(lambda x: remove_after_digits_keep_separators(str(x)))
            df_2025.iloc[:, 5] = df_2025.iloc[:, 5].apply(lambda x: convert_to_uppercase(str(x)))
            df_2025.iloc[:, 5] = df_2025.iloc[:, 5].apply(lambda x: remove_characters_around_hyphen_regex(str(x)))
            # 替代部分不规范字符
            df_2025.replace({
                'LongVSPI': '0flash',
                'LongVI2C': '1flash',
                'Longv0flash': '0flash',
                'longv1flash': '1flash',
                'longv0flashSPI': '0flash',
                'I2C': '1flash',
                'longv1flashI2C':'1flash',
                'longH1flashI2C':'0flash',
                'longH0flashSPI':'0flash',
                'Longv1Flash':'1flash',
                'Longv1FlashI2C':'1flash',
                'LongvWithFlashI2C':'1flash',
                'longv0flash':'0flash',
                'LongV1FlashI2C':'1flash',
                'COST6.56': 'CSOT6.56',
                '否': '关',
                '\\':'NULL',
            }, inplace=True)
            df_2025 = df_2025.fillna('NULL')
            # 删除字符串中的空格和换行符
            df_2025 = df_2025.replace({'\n': '', ' ': ''}, regex=True)
            # df_2025.iloc[:, 14].to_csv('CLASS_2025.txt', index=False, header=False, encoding='utf-8')

            #2026年度excel
            df_2026 = pd.read_excel(excel_file, sheet_name='2026年TDDI', skiprows=[0])
            df_2026 = pd.DataFrame(df_2026)
            # # 删除字符串中的空格和换行符
            df_2026 = df_2026.replace({'\n': '', ' ': ''}, regex=True)

            # 去除模组厂字符串
            df_2026.iloc[:, 3] = df_2026.iloc[:, 3].apply(lambda x: remove_english_keep_chinese(str(x)))
            df_2026.iloc[:, 3] = df_2026.iloc[:, 3].apply(lambda x: remove_parentheses(str(x)))
            # 去除玻璃后面型号备注
            df_2026.iloc[:, 5] = df_2026.iloc[:, 5].apply(lambda x: remove_after_digits_keep_separators(str(x)))
            df_2026.iloc[:, 5] = df_2026.iloc[:, 5].apply(lambda x: convert_to_uppercase(str(x)))
            df_2026.iloc[:, 5] = df_2026.iloc[:, 5].apply(lambda x: remove_characters_around_hyphen_regex(str(x)))
            # 替代部分不规范字符
            df_2026.replace({
                'LongVSPI': '0flash',
                'LongVI2C': '1flash',
                'Longv0flash': '0flash',
                'longv1flash': '1flash',
                'longv0flashSPI': '0flash',
                'I2C': '1flash',
                'longv1flashI2C':'1flash',
                'longH1flashI2C':'0flash',
                'longH0flashSPI':'0flash',
                'Longv1Flash':'1flash',
                'Longv1FlashI2C':'1flash',
                'LongvWithFlashI2C':'1flash',
                'longv0flash':'0flash',
                'LongV1FlashI2C':'1flash',
                'COST6.56': 'CSOT6.56',
                '否':'关',
                '\\': 'NULL',

            }, inplace=True)
            df_2026 = df_2026.fillna('NULL')
            # 删除字符串中的空格和换行符
            df_2026 = df_2026.replace({'\n': '', ' ': ''}, regex=True)
            # df_2026.iloc[:, 14].to_csv('CLASS_2026.txt', index=False, header=False, encoding='utf-8')
        df_combined = pd.concat([df_2025, df_2026], ignore_index=True)
        # df_combined.iloc[:, 14].to_csv('CLASS_df_combined.txt', index=False, header=False, encoding='utf-8')
        return df_combined

    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None

# debug funciton
# test_cases = [
#     "信雅达S200X",
#     "晶泰JTZ068062X0B",
#     "德智欣Y92372",
# ]
# for test in test_cases:
#     result = remove_english_keep_chinese(test)
#     print(f"原始: {test}")
#     print(f"结果: {result}")
#     print("-" * 30)

# file_name = 'LCD TDDI华南项目汇总_0122.xlsx'
# sheet_name = '2025年TDDI'
# sheet_name1 = '2026年TDDI'
# df= read_excel_for_analyseHN(file_name)
# df_2026= read_excel_for_analyseHN(file_name,sheet_name1)
# print(df_2025.iloc[:, 3])
# print(df_2025.iloc[:, 4])
# print(df_2025.iloc[:, 5])
# 合并获取第5列数据并去重
# df_combined = pd.concat([df_2025, df_2026], ignore_index=True)
# print(df.iloc[:, 3])
# unique_values = df.iloc[:, 3].drop_duplicates()
# unique_values_sorted = unique_values.sort_values(ascending=True)
# df.iloc[:, 3].to_csv('模组厂.txt', index=False, header=False, encoding='utf-8')
# unique_values_sorted.to_csv('模组厂不重.txt', index=False, header=False, encoding='utf-8')
# print(unique_values)
# print("unique_values已保存到 glass.txt 文件")
# df_combined.iloc[:,3].drop_duplicates().to_csv('module.txt', index=False, header=False, encoding='utf-8')