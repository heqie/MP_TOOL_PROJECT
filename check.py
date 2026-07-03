#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT 
@File    ：check.py
@Author  ：zxh
@Date    ：2025/3/19 16:21 
'''
import pandas as pd
import re
import openpyxl


def check_sequence(df):
    """
    基于C列项目号检查顺序排列
    :param df: 处理过的二维表格
    :return: 错误值
    """
    errors = []
    cleaned_project_numbers = df.iloc[:, 2]
    # 初始化序号
    expected_sequence = 1

    for idx, (project_number, sequence) in enumerate(zip(cleaned_project_numbers, df.iloc[:, 0])):
        if pd.isna(sequence) or sequence == '' or sequence == 0:
            errors.append((idx + 2, '序号', "序号为空值"))
            if idx > 0 and project_number != cleaned_project_numbers[idx - 1]:
                expected_sequence += 1
            continue

        # 如果当前项目号与前一个项目号不同，则序号应递增
        if idx > 0 and project_number != cleaned_project_numbers[idx - 1]:
            expected_sequence += 1

        # 检查序号是否正确，确保将sequence转换为整数
        try:
            sequence_int = int(sequence)
        except ValueError:
            errors.append((idx + 2, '序号', f"序号 '{sequence}' 不是有效的整数"))
            continue

        # 检查序号是否正确
        if sequence_int != expected_sequence:
            errors.append((idx + 2, '序号', f"期望 {expected_sequence}, 实际 {sequence_int}"))

    return errors


def check_sequence_for_fix(df):
    """
    基于C列项目号检查顺序排列，并只返回相同项目号且相同错误的第一行错误，用于校正序号
    :param df: 处理过的二维表格
    :return: 错误值列表，每个错误值为 (行号, 列名, 错误信息)
    """
    errors = []
    cleaned_project_numbers = df.iloc[:, 2]
    expected_sequence = 1  # 初始化期望的序号

    # 用于存储上一个错误信息和项目号
    last_error = None
    last_project_number = None

    for idx, (project_number, sequence) in enumerate(zip(cleaned_project_numbers, df.iloc[:, 0])):
        # 如果当前项目号与前一个项目号不同，则期望序号递增
        if idx > 0 and project_number != cleaned_project_numbers[idx - 1]:
            expected_sequence += 1

        # 检查当前行的序号是否为空值或无效值
        if pd.isna(sequence) or sequence == '' or sequence == 0:
            # 如果当前项目号与上一个项目号相同，且错误信息相同，则跳过
            if project_number == last_project_number and current_error == last_error:
                continue
            current_error = (idx + 2, '序号', expected_sequence, "序号为空值")
            # 否则，记录当前错误
            errors.append(current_error)
            last_error = current_error
            last_project_number = project_number
            continue

        # 检查序号是否为有效整数
        try:
            sequence_int = int(sequence)
        except ValueError:
            # 如果当前项目号与上一个项目号相同，且错误信息相同，则跳过
            if project_number == last_project_number and current_error == last_error:
                continue
            current_error = (idx + 2, '序号', f"序号 '{sequence}' 不是有效的整数")
            # 否则，记录当前错误
            errors.append(current_error)
            last_error = current_error
            last_project_number = project_number
            continue

        # 检查序号是否正确
        if sequence_int != expected_sequence:
            # 如果当前项目号与上一个项目号相同，且错误信息相同，则跳过
            if project_number == last_project_number and current_error == last_error:
                continue
            current_error = (idx + 2, '序号', expected_sequence, sequence_int)
            # 否则，记录当前错误
            errors.append(current_error)
            last_error = current_error
            last_project_number = project_number

    return errors


def check_grade(column):
    """
    检查等级
    :param column:list中等级所在的那一列数据
    :return: 错误值
    """
    valid_chars = {'-', 'A', 'V', 'v', 'G'}
    errors = []
    for idx, value in enumerate(column):
        if pd.isna(value):
            continue
        if value == 0:
            continue
        if value not in valid_chars:
            errors.append((idx + 2, '等级', value))
    return errors


def check_ic_model(column, valid_IC):
    """
    检查IC型号,对比compare文件列表
    :param column: list中IC所在的那一列数据
    :param valid_IC: 有效IC的对比
    :return: 错误值
    """
    errors = []
    for idx, value in enumerate(column):
        if pd.isna(value) or value == 0:
            errors.append((idx + 2, 'IC', "IC为空值"))
            continue
        if value == 7272:
            continue
        if str(value) not in valid_IC:  # 将值转换为字符串进行比较
            errors.append((idx + 2, 'IC', value))
    return errors


def check_glass_model(column, valid_models):
    """
    检查玻璃型号，对比compare文件列表
    :param column: list中玻璃所在的那一列数据
    :param valid_models: 有效玻璃型号的对比
    :return: 错误值
    """
    errors = []
    for idx, value in enumerate(column):
        if pd.isna(value) or value == 0:
            errors.append((idx + 2, '玻璃', "玻璃为空值"))
            continue
        if value not in valid_models:
            errors.append((idx + 2, '玻璃', value))
    return errors


def check_interface(column):
    """
    检查接口定义,'SPI', 'I2C'两种
    :param column: list中接口所在的那一列数据
    :return: 错误值
    """
    valid_interfaces = {'SPI', 'I2C'}
    errors = []
    for idx, value in enumerate(column):
        if pd.isna(value) or value == 0:
            errors.append((idx + 2, '接口', "接口为空值"))
            continue
        if value not in valid_interfaces:
            errors.append((idx + 2, '接口', value))
    return errors


def check_flash(column, interface_column):
    """
    检查flash，对比接口即SPI->N,I2C->Y
    :param column: list中flash所在的那一列数据
    :param interface_column: list中接口所在的那一列数据
    :return:错误值
    """
    errors = []
    for idx, (interface, flash) in enumerate(zip(interface_column, column)):
        if pd.isna(interface) or pd.isna(flash) or interface == 0 or flash == 0:
            if interface == 0 and flash != 0:
                errors.append((idx + 2, 'flash', "接口为空值，无法判断"))
            if flash == 0:
                errors.append((idx + 2, 'flash', "flash为空值"))
            continue
        if not ((interface == 'SPI' and flash == 'N') or (interface == 'I2C' and flash == 'Y')):
            errors.append((idx + 2, 'flash', flash))
    return errors


def check_tool_version(df, valid_version):
    """
    检查TOOL版本,对比compare文件列表
    :param df:list中TOOL版本所在的那一列数据
    :param valid_version:有效版本对比
    :return:错误值
    """
    update_column = df.iloc[:, 15]
    site_column = df.iloc[:, 10]
    column = df.iloc[:, 13]
    errors = []
    for idx, (value, update_value, site_value) in enumerate(zip(column, update_column, site_column)):
        # 如果站点为 "整机"，则 TOOL 版本忽略
        if site_value == "整机":
            continue
        # 如果站点为模组，则判断
        elif site_value == "模组":
            # 如果更新说明有字符，则 TOOL 版本不能为空
            if pd.notna(update_value) and str(update_value).strip() != '' and update_value != 0:
                if pd.isna(value) or value == 0 or value == '':
                    errors.append((idx + 2, 'Tool版本号', "TOOL版本为空值（有更新说明）"))
                elif value not in valid_version:
                    errors.append((idx + 2, 'Tool版本号', value))
            # 如果更新说明为空，则 TOOL 版本可以为空
            else:
                if pd.notna(value) and value != 0 and value != '' and value not in valid_version:
                    errors.append((idx + 2, 'Tool版本号', value))
    return errors


def check_date_format(df):
    """
    检查发布日期格式,YYY.MM.DD或YYYY.M.D
    :param df:list中发布日期所在的那一列数据
    :return:错误值
    """
    pattern = re.compile(r'^\d{4}\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])$')
    errors = []
    column = df.iloc[:, 16]
    update_column = df.iloc[:, 15]
    for idx, (value, update_value) in enumerate(zip(column, update_column)):
        # 如果更新说明有字符，则发布日期不能为空
        if pd.notna(update_value) and str(update_value).strip() != '' and update_value != 0:
            if pd.isna(value) or value == 0 or value == '':
                errors.append((idx + 2, '发布日期', "发布日期为空值（有更新说明）"))
            elif not pattern.match(str(value)):
                errors.append((idx + 2, '发布日期', value))
        # 如果更新说明为空，则发布日期可以为空
        else:
            if pd.notna(value) and value != 0 and value != '' and not pattern.match(str(value)):
                errors.append((idx + 2, '发布日期', value))
    return errors


def fix_sequence(file_path, errors):
    """
    用于序列修改
    :param file_path: 文件路径
    :param errors: 需要修改的数值信息，该数值取自check_sequence_for_fix(df)
    :return:
    """
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    # 指定列（例如：B列，列号为2）
    target_column = 1
    # 获取目标单元格
    for idx in errors:
        target_row = idx[0]
        new_value = idx[2]

        target_cell = sheet.cell(row=target_row, column=target_column)

        # 检查目标单元格是否在合并单元格中
        is_merged = False
        for merged_range in sheet.merged_cells.ranges:
            # 检查目标单元格是否在当前合并范围内
            if (merged_range.min_row <= target_row <= merged_range.max_row and
                    merged_range.min_col <= target_column <= merged_range.max_col):
                is_merged = True
                # 找到合并区域的左上角单元格
                top_left_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                # 修改左上角单元格的值
                top_left_cell.value = new_value
                break

        # 如果目标单元格不在合并单元格中，直接替换该单元格的值
        if not is_merged:
            target_cell.value = new_value

    # return workbook
    # # 保存文件
    output_file_path = 'modified_excel_file.xlsx'
    workbook.save(output_file_path)

    # print(f"文件已保存为 {output_file_path}")