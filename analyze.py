#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT 
@File    ：analyze.py
@Author  ：zxh
@Date    ：2025/3/19 16:22 
'''
import pandas as pd
from datetime import datetime
import re

# 全局变量
IC_TYPES = ['7272', '7202M', '7202H']
FACTORIES = ['海菲', '信利', '易快来', '联创', '创维', '同兴达', '海盛捷', '壹星', '精卓', '三龙',
             '合力泰', '华显', '维立', '亿华', '立德', '晶泰', '德智欣', '中光电', '沛宏',
             '欣欣光电', '众铭安', '天正达', '瑞恒光电', '清创高', '汉龙时代', '晶胜通',
             '龙煜', '金宏光电', '威达光电', '惠科', '华视', '正金晶光电', '华映', '长信新显',
             '宏凯', '泰启', '百业', '共赢', '德实', '京龙', '如新电子', '皓显', '大通显示',
             '高展', '康华显通', '煜鑫', '宏利超显', '菲触显视', '轩达', '钜沣', '鹰芒技术', '德普特', '元格', '亿普拉斯','万联','重联']
GLASS_TYPES = ['CSOT', 'TM', 'TRULY', 'BOE', 'CTO', 'PANDA',
               'SHARP', 'MDT', 'HKC', 'HSD', 'INX', 'IVO', 'CTC']

TARGET_ICS = ['7272', '7202M', '7202H']
YEARS = ['2022年度', '2023年度', '2024年度', '2025年度']


def statistic_project_status(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL'):
    """
    根据条件筛选更新次数的项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :return: 项目名，统计结果
    """
    # 数据预处理
    df_filtered = df[df.iloc[:, 10] == '模组'].copy()

    # 标准化各列数据
    def normalize_ic(ic):
        ic = str(ic).upper()
        if '7202M' in ic: return '7202M'
        if '7202H' in ic: return '7202H'
        return ic if ic in IC_TYPES else 'OTHER'

    def normalize_glass(g):
        match = re.match(r'^([A-Za-z]+)', str(g))
        return match.group(1).upper() if match else str(g).upper()

    df_filtered.loc[:, 'IC_Norm'] = df_filtered.iloc[:, 6].apply(normalize_ic)
    df_filtered.loc[:, 'Factory'] = df_filtered.iloc[:, 5].astype(str)
    df_filtered.loc[:, 'Glass_Norm'] = df_filtered.iloc[:, 7].apply(normalize_glass)
    df_filtered.loc[:, 'Flash'] = df_filtered.iloc[:, 9].astype(str).str.upper().apply(
        lambda x: 'Y' if x.startswith('Y') else 'N')
    df_filtered.loc[:, 'Year'] = df_filtered.iloc[:, 16].astype(str).apply(
        lambda x: f"{datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year}"
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x)
        else 'UNKNOWN')

    # 创建唯一项目ID
    df_filtered.loc[:, 'Project_ID'] = (
            df_filtered.iloc[:, 2].astype(str) + '_' +
            df_filtered.iloc[:, 5].astype(str) + '_' +
            df_filtered.iloc[:, 7].astype(str)
    )

    # 应用筛选条件
    if ic_type != 'ALL':
        if ic_type == '7202M':
            df_filtered = df_filtered[df_filtered['IC_Norm'] == '7202M']
        elif ic_type == '7202H':
            df_filtered = df_filtered[df_filtered['IC_Norm'] == '7202H']
        else:
            df_filtered = df_filtered[df_filtered['IC_Norm'] == ic_type]

    if factory != 'ALL':
        df_filtered = df_filtered[df_filtered['Factory'] == factory]

    if glass != 'ALL':
        df_filtered = df_filtered[df_filtered['Glass_Norm'] == glass.upper()]

    if flash != 'ALL':
        df_filtered = df_filtered[df_filtered['Flash'] == flash.upper()]

    if year != 'ALL':
        df_filtered = df_filtered[df_filtered['Year'] == year]

    # 统计更新次数（P列）
    df_filtered.loc[:, 'Update_Count'] = df_filtered.iloc[:, 15].astype(str).apply(
        lambda x: 0 if x == '0' else 1)

    # 分组统计
    project_stats = df_filtered.groupby('Project_ID').agg({
        'Update_Count': 'sum',
        'IC_Norm': 'first',
        'Factory': 'first',
        'Glass_Norm': 'first',
        'Flash': 'first',
        'Year': 'first'
    }).reset_index()

    # 生成分布统计
    dist_stats = project_stats['Update_Count'].value_counts().sort_index().reset_index()
    dist_stats.columns = ['Update_Count', 'Project_Count']

    # 打印筛选摘要
    print("▌筛选条件汇总")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {len(project_stats)}）")
    # print(dist_stats.to_string(index=False))
    print("─" * 40)

    conditions = []
    if ic_type != 'ALL': conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': conditions.append(f"Flash: {flash}")
    if year != 'ALL': conditions.append(f"年度: {year}")

    if len(project_stats) == 0:
        print("没有找到符合条件的项目")
        return conditions, pd.Series(dtype=int)  # 返回空Series

    return conditions, dist_stats


def statistic_ic_projects(df, factory='ALL', glass='ALL', flash='ALL', year='ALL'):
    """
    根据条件筛选IC型号项目数量
    :param df: 传入二维数组
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :return: 项目名，统计结果
    """
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Factory'] = df.iloc[:, 5].astype(str)  # F列
    df['Glass_before'] = df.iloc[:, 7].astype(str)
    df['Glass'] = df.iloc[:, 7].astype(str).apply(  # H列
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper() if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    df['IC_Type'] = df.iloc[:, 6].astype(str)  # G列
    df['Flash'] = df.iloc[:, 9].astype(str).apply(  # J列
        lambda x: 'Y' if str(x).upper().startswith('Y') else 'N'
    )
    df['Year'] = df.iloc[:, 16].astype(str).apply(  # Q列
        lambda x: str(datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year)
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知'
    )

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]
    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]
    if year != 'ALL':
        df = df[df['Year'] == year]

    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['Project_Name'] + '_' + df['Factory'] + '_' + df['Glass_before']

    # IC型号匹配（7202M/7202H模糊匹配）
    pattern = re.compile(r'7272|7202[MH]')
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
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(ic_counts.sum())}）")
    print("─" * 40)

    title_conditions = []
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")

    if len(ic_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, ic_counts


def statistic_module_factory(df, ic_type='ALL', glass='ALL', flash='ALL', year='ALL'):
    """
    根据条件筛选模组厂项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :return: 项目名，统计结果
    """
    df = df.copy()
    # 标准化各列数据
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Factory'] = df.iloc[:, 5].astype(str)  # F列
    df['Glass_before'] = df.iloc[:, 7].astype(str)
    df['Glass'] = df.iloc[:, 7].astype(str).apply(  # H列
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper()
        if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    df['IC_Type'] = df.iloc[:, 6].astype(str)  # G列
    df['Flash'] = df.iloc[:, 9].astype(str).apply(  # J列
        lambda x: 'Y' if str(x).upper().startswith('Y') else 'N'
    )
    df['Year'] = df.iloc[:, 16].astype(str).apply(  # Q列
        lambda x: str(datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year)
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知'
    )

    # 应用筛选条件
    if ic_type != 'ALL':
        if ic_type == '7202M':
            df = df[df['IC_Type'].str.contains('7202M', na=False)]
        elif ic_type == '7202H':
            df = df[df['IC_Type'].str.contains('7202H', na=False)]
        else:
            df = df[df['IC_Type'] == ic_type]

    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]

    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]

    if year != 'ALL':
        df = df[df['Year'] == year]

    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['Project_Name'] + '_' + df['Factory'] + '_' + df['Glass_before']

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
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(factory_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")

    if len(factory_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, factory_counts


def statistic_glass_projects(df, ic_type='ALL', factory='ALL', flash='ALL', year='ALL'):
    """
    根据条件筛选玻璃厂项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :return: 项目名，统计结果
    """
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Factory'] = df.iloc[:, 5].astype(str)  # F列
    df['Glass_before'] = df.iloc[:, 7].astype(str)
    df['Glass'] = df.iloc[:, 7].astype(str).apply(  # H列
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper()
        if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    df['IC_Type'] = df.iloc[:, 6].astype(str)  # G列
    df['Flash'] = df.iloc[:, 9].astype(str).apply(  # J列
        lambda x: 'Y' if str(x).upper().startswith('Y') else 'N'
    )
    df['Year'] = df.iloc[:, 16].astype(str).apply(  # Q列
        lambda x: str(datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year)
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知'
    )

    # 应用筛选条件
    if ic_type != 'ALL':
        if ic_type == '7202M':
            df = df[df['IC_Type'].str.contains('7202M', na=False)]
        elif ic_type == '7202H':
            df = df[df['IC_Type'].str.contains('7202H', na=False)]
        else:
            df = df[df['IC_Type'] == ic_type]

    if factory != 'ALL':
        df = df[df['Factory'] == factory]

    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]

    if year != 'ALL':
        df = df[df['Year'] == year]

    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['Project_Name'] + '_' + df['Factory'] + '_' + df['Glass_before']

    # 只保留定义的玻璃类型
    df = df[df['Glass'].isin(GLASS_TYPES)]

    # 统计各玻璃类型的项目数量（按唯一ID去重）
    glass_counts = df.drop_duplicates('Unique_ID')['Glass'].value_counts()

    # 过滤掉数量为0的玻璃类型并按数量排序
    glass_counts = glass_counts[glass_counts > 0].sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(glass_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")

    if len(glass_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, glass_counts


def statistic_flash_projects(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL'):
    """
    根据条件筛选Flash项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param year: 2023/2024/2025/ALL
    :return: 项目名，统计结果
    """
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Factory'] = df.iloc[:, 5].astype(str)  # F列
    df['Glass_before'] = df.iloc[:, 7].astype(str)
    df['Glass'] = df.iloc[:, 7].astype(str).apply(  # H列
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper()
        if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    df['IC_Type'] = df.iloc[:, 6].astype(str)  # G列
    df['Flash'] = df.iloc[:, 9].astype(str).apply(  # J列
        lambda x: 'Y' if str(x).upper().startswith('Y') else 'N'
    )
    df['Year'] = df.iloc[:, 16].astype(str).apply(  # Q列
        lambda x: str(datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year)
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知'
    )

    # 应用筛选条件
    if ic_type != 'ALL':
        if ic_type == '7202M':
            df = df[df['IC_Type'].str.contains('7202M', na=False)]
        elif ic_type == '7202H':
            df = df[df['IC_Type'].str.contains('7202H', na=False)]
        else:
            df = df[df['IC_Type'] == ic_type]

    if factory != 'ALL':
        df = df[df['Factory'] == factory]

    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]

    if year != 'ALL':
        df = df[df['Year'] == year]

    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['Project_Name'] + '_' + df['Factory'] + '_' + df['Glass_before']

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计Flash状态
    flash_counts = unique_projects['Flash'].value_counts()
    flash_counts = flash_counts.reindex(['Y', 'N'], fill_value=0)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(flash_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL': title_conditions.append(f"年度: {year}")

    if len(flash_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series
    return title_conditions, flash_counts


def statistic_project_by_year(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL'):
    """
    根据条件筛选年度项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :return: 项目名，统计结果
    """
    # 数据预处理
    df = df.copy()
    # 标准化各列数据
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Factory'] = df.iloc[:, 5].astype(str)  # F列
    df['Glass_before'] = df.iloc[:, 7].astype(str)
    df['Glass'] = df.iloc[:, 7].astype(str).apply(  # H列
        lambda x: re.match(r'^([A-Za-z]+)', str(x)).group(1).upper()
        if re.match(r'^([A-Za-z]+)', str(x)) else x.upper()
    )
    df['IC_Type'] = df.iloc[:, 6].astype(str)  # G列
    df['Flash'] = df.iloc[:, 9].astype(str).apply(  # J列
        lambda x: 'Y' if str(x).upper().startswith('Y') else 'N'
    )
    df['Year'] = df.iloc[:, 16].astype(str).apply(  # Q列
        lambda x: f"{datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year}年度"
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知年度'
    )

    # 应用筛选条件
    if ic_type != 'ALL':
        if ic_type == '7202M':
            df = df[df['IC_Type'].str.contains('7202M', na=False)]
        elif ic_type == '7202H':
            df = df[df['IC_Type'].str.contains('7202H', na=False)]
        else:
            df = df[df['IC_Type'] == ic_type]

    if factory != 'ALL':
        df = df[df['Factory'] == factory]

    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]

    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]

    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['Project_Name'] + '_' + df['Factory'] + '_' + df['Glass_before']

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    year_counts = unique_projects['Year'].value_counts()
    year_counts = year_counts.reindex(YEARS, fill_value=0)
    # print(year_counts)
    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(year_counts.sum())}）")
    print("─" * 40)


    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")

    if len(year_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, year_counts
