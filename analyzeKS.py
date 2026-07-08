#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT
@File    ：analyzeKS.py
@Author  ：zxh
@Date    ：2025/4/28 11:22 
'''

import pandas as pd
from datetime import datetime
import re
import matplotlib.pyplot as plt
import seaborn as sns
from analyze import *


plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

KS_TYPES = datas.iloc[:, 5].dropna().tolist()
PRINCIPAL = datas.iloc[:, 6].dropna().tolist()
# print(KS_TYPES)
# print(PRINCIPAL)
# KS_TYPES = ['DPcode', 'FW效果', 'IC来料', '环境干扰', '模组工艺', '人为因素', '玻璃', '误判']
# PRINCIPAL = ['黄耀', '吴宏博', '丁淑芳', '成鹏', '李尚聪', '卓秀慧']


# debug
# df = pd.read_excel('Project MP List&MP Flow List&CUT4兼容&HV noise (12).xlsx',sheet_name='客诉记录')
# df = read_excel_for_analyseKS('Project MP List&MP Flow List&CUT4兼容&HV noise (12).xlsx',sheet_name='客诉记录')
# # df.to_excel('output.xlsx', index=False)
# print(df.iloc[32,12])
# print(df.iloc[:,0])
# # 先尝试解析文本日期
# df['date_column'] = pd.to_datetime(df.iloc[:,5], format='%Y年%m月%d日', errors='coerce')
#
# # 剩下的解析为 Excel 数值日期
# mask = df['date_column'].isna()
# df.loc[mask, 'date_column'] = pd.to_datetime(
#     pd.to_numeric(df.iloc[:,5][mask], errors='coerce'),
#     unit='D',
#     origin='1899-12-30'
# )
# print(df['date_column'])

def parse_mixed_date_column_to_year(analyse_type, data, column, date_format='%Y年%m月%d日', excel_origin='1899-12-30'):
    """
    解析混合格式的日期列，并返回年份  日期格式只能是%Y年%m月%d日文本内容或excel公式内容两种
    :param data: 原始二维数组
    :param column:对应列
    :param date_format: 文本日期格式
    :param excel_origin: excel表公式日期基准
    :param default_year: int/None，无法解析时的默认值（默认None，返回NaN）
    :return:
    """
    # 提取目标列数据
    if isinstance(column, int):
        date_series = data.iloc[:, column]
    else:
        date_series = data[column]

    # 解析文本格式日期
    parsed_dates = pd.to_datetime(date_series, format=date_format, errors='coerce')

    # 剩余未解析作为Excel数值日期处理
    mask = parsed_dates.isna()
    if mask.any():
        numeric_dates = pd.to_numeric(date_series[mask], errors='coerce')
        parsed_dates[mask] = pd.to_datetime(
            numeric_dates,
            unit='D',
            origin=excel_origin
        )

    # 提取年份，无法解析的返回NaN
    years = parsed_dates.dt.year
    if analyse_type == '年度':
        formatted_years = years.apply(lambda x: f"{int(x)}年度" if pd.notna(x) else None)
    else:
        formatted_years = years.apply(lambda x: f"{int(x)}年度" if pd.notna(x) else None)

    return parsed_dates, formatted_years


def normolization_dataKS(analyse_type, df):
    '''
    规范化excel表客诉记录数据结构
    :param df: 二维数组
    :return: 规范后的数据
    '''
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['sequence'] = df.iloc[:, 0].astype(str)  # 序号
    df['Factory'] = df.iloc[:, 1].astype(str)  # 模组厂
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # 项目号
    df['IC_Type'] = df.iloc[:, 3].astype(str)  # IC
    df['Glass_before'] = df.iloc[:, 4].astype(str)  # 玻璃
    df['Glass'] = df.iloc[:, 4].astype(str).apply(  # 将玻璃格式化
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper() if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    parsed_dates, formatted_years = parse_mixed_date_column_to_year(analyse_type, df, 5)
    df['Date'] = parsed_dates.astype(str)
    df['Year'] = formatted_years.astype(str)  # 提取开始的年份
    finish_dates, finish_years = parse_mixed_date_column_to_year(analyse_type, df, 6)
    df['finish_Date'] = finish_dates.astype(str)
    # print(df['Year'])
    df['Transfer_AE'] = df.iloc[:, 10].astype(str)  # 转接AE与否
    df['Over'] = df.iloc[:, 11].astype(str)  # 结案与否
    df['Type'] = df.iloc[:, 12].astype(str)  # 客诉类型
    df['Principal'] = df.iloc[:, 13].astype(str)  # 负责人
    # 创建唯一项目ID（序号+项目+模组厂+玻璃）
    df['Unique_ID'] = df['sequence'] + '_' + df['Project_Name'] + '_' + df['Factory'] + '_' +df['IC_Type']+'_' + df['Glass_before'] + '_' + \
                      df['Date'] + '_' + df['finish_Date']

    return df


def statistic_ic_projects_KS(df, factory='ALL', glass='ALL', year='ALL', transfer_AE='ALL', over='ALL', type='ALL',
                             principal='ALL'):
    """
    IC基础为X轴进行筛选统计
    :param df: 传入二维数组
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param year: 2023/2024/2025/ALL
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """

    # 标准化各列数据
    df = normolization_dataKS('ic型号', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # IC型号匹配（7202M/7202H模糊匹配）
    pattern = re.compile(r'7202MA|7272CA|7272|7202[MH]|7302')
    df['Matched_IC'] = df['IC_Type'].apply(
        lambda x: pattern.search(x).group() if pattern.search(x) else None
    )

    # 过滤目标IC型号
    filtered_df = df[df['Matched_IC'].isin(TARGET_ICS)]

    # 按唯一ID去重后统计
    ic_counts = filtered_df.drop_duplicates('Unique_ID')['Matched_IC'].value_counts()
    ic_counts = ic_counts.reindex(TARGET_ICS, fill_value=0)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(ic_counts.sum())}）")
    print("─" * 40)

    title_conditions = []
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(ic_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, ic_counts


def statistic_module_factory_KS(df, ic_type='ALL', glass='ALL', year='ALL', transfer_AE='ALL', over='ALL', type='ALL',
                                principal='ALL'):
    """
    模组厂为基础进行筛选
    :param df: 传入二维数组
    :param ic_type: IC型号
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param year: 2023/2024/2025/ALL
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('模组厂', df)

    # 应用筛选条件
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 只保留定义的模组厂
    df = df[df['Factory'].isin(FACTORIES)]

    # 统计各模组厂的项目数量（按唯一ID去重）
    factory_counts = df.drop_duplicates('Unique_ID')['Factory'].value_counts()

    # print(factory_counts)
    # 过滤掉数量为0的模组厂并按数量排序
    factory_counts = factory_counts[factory_counts > 0].sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(factory_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC型号: {ic_type}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(factory_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, factory_counts


def statistic_glass_projects_KS(df, factory='ALL', ic_type='ALL', year='ALL', transfer_AE='ALL', over='ALL', type='ALL',
                                principal='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param year: 2023/2024/2025/ALL
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('玻璃厂', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 只保留定义的玻璃类型
    df = df[df['Glass'].isin(GLASS_TYPES)]

    # 统计各玻璃类型的项目数量（按唯一ID去重）
    glass_counts = df.drop_duplicates('Unique_ID')['Glass'].value_counts()

    # 过滤掉数量为0的玻璃类型并按数量排序
    glass_counts = glass_counts[glass_counts > 0].sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(glass_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if ic_type != 'ALL': title_conditions.append(f"IC型号: {ic_type}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    return title_conditions, glass_counts


def statistic_project_by_year_KS(df, factory='ALL', ic_type='ALL', glass='ALL', transfer_AE='ALL', over='ALL',
                                 type='ALL', principal='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param glass: 玻璃型号
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('年度', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    year_counts = unique_projects['Year'].value_counts()
    year_counts = year_counts.reindex(YEARS, fill_value=0)
    # print(year_counts)
    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(year_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(year_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, year_counts


def statistic_project_transferAE_KS(df, factory='ALL', ic_type='ALL', glass='ALL', year='ALL', over='ALL', type='ALL',
                                    principal='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param glass: 玻璃型号
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('转接AE与否', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    transferAE_counts = unique_projects['Transfer_AE'].value_counts()
    transferAE_counts = transferAE_counts.reindex(['是', '否'], fill_value=0)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(transferAE_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(transferAE_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, transferAE_counts


def statistic_project_over_KS(df, factory='ALL', ic_type='ALL', glass='ALL', year='ALL', transfer_AE='ALL', type='ALL',
                              principal='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param glass: 玻璃型号
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('结案与否', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if type != 'ALL':
        df = df[df['Type'] == type]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    over_counts = unique_projects['Over'].value_counts()
    over_counts = over_counts.reindex(['是', '否'], fill_value=0)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(over_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(over_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, over_counts


def statistic_project_type_KS(df, factory='ALL', ic_type='ALL', glass='ALL', year='ALL', transfer_AE='ALL', over='ALL',
                              principal='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param glass: 玻璃型号
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('客诉类型', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if principal != 'ALL':
        df = df[df['Principal'] == principal]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    type_counts = unique_projects['Type'].value_counts()
    type_counts = type_counts.reindex(KS_TYPES, fill_value=0)

    type_counts = type_counts.sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"负责人: {principal if principal != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(type_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if principal != 'ALL': title_conditions.append(f"负责人: {principal}")

    if len(type_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, type_counts


def statistic_project_principal_KS(df, factory='ALL', ic_type='ALL', glass='ALL', year='ALL', transfer_AE='ALL',
                                   over='ALL', type='ALL'):
    """
    玻璃厂为基础进行筛选
    :param df: 传入二维数组
    :param factory:模组厂
    :param ic_type: IC型号
    :param glass: 玻璃型号
    :param transfer_AE:是否转接AE(是/否/ALL)
    :param over:是否结案(是/否/ALL)
    :param type:客诉类型(ALL/DP code/FW效果/IC来料/环境干扰/模组工艺/人为因素/玻璃来料)
    :param principal:负责人
    :return: 项目名，统计结果
    """
    # 标准化各列数据
    df = normolization_dataKS('负责人', df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if transfer_AE != 'ALL':
        df = df[df['Transfer_AE'] == transfer_AE]
    if over != 'ALL':
        df = df[df['Over'] == over]
    if type != 'ALL':
        df = df[df['Type'] == type]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    principal_counts = unique_projects['Principal'].value_counts()
    principal_counts = principal_counts.reindex(PRINCIPAL, fill_value=0)

    principal_counts = principal_counts.sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"转接AE与否: {transfer_AE if transfer_AE != 'ALL' else '全部'}")
    print(f"结案与否: {over if over != 'ALL' else '全部'}")
    print(f"客诉类型: {type if type != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(principal_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")
    if transfer_AE != 'ALL': title_conditions.append(f"转接AE与否: {transfer_AE}")
    if over != 'ALL': title_conditions.append(f"结案与否: {over}")
    if type != 'ALL': title_conditions.append(f"客诉类型: {type}")

    if len(principal_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, principal_counts
