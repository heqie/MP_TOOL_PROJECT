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
from read_excel import *

# # 全局变量
data_file = 'compare_information.xlsx'
datas = read_compare_excel(data_file, sheet_names='Sheet2')
TARGET_ICS = datas.iloc[:, 0].dropna().astype(str).tolist()
FACTORIES = datas.iloc[:, 1].dropna().tolist()
GLASS_TYPES = datas.iloc[:, 2].dropna().tolist()
YEARS = datas.iloc[:, 3].dropna().tolist()
GRADES = datas.iloc[:, 4].dropna().tolist()
GRADES.append('NULL')
PUBLISHER = datas.iloc[:, 7].dropna().tolist()

# print(TARGET_ICS)
# print(FACTORIES)
# print(GLASS_TYPES)
# print(YEARS)
# print(GRADES)
# IC_TYPES = ['7272', '7202M', '7202H']
# FACTORIES = ['海菲', '信利', '易快来', '联创', '创维', '同兴达', '海盛捷', '壹星', '精卓', '三龙',
#              '合力泰', '华显', '维立', '亿华', '立德', '晶泰', '德智欣', '中光电', '沛宏',
#              '欣欣光电', '众铭安', '天正达', '瑞恒光电', '清创高', '汉龙时代', '晶胜通',
#              '龙煜', '金宏光电', '威达光电', '惠科', '华视', '正金晶光电', '华映', '长信新显',
#              '宏凯', '泰启', '百业', '共赢', '德实', '京龙', '如新电子', '皓显', '大通显示',
#              '高展', '康华显通', '煜鑫', '宏利超显', '菲触显视', '轩达', '钜沣', '鹰芒技术',
#              '德普特', '元格', '亿普拉斯', '万联', '重联', '日日佳', '中正威', '天山电子', '凯晟']
# GLASS_TYPES = ['CSOT', 'TM', 'TRULY', 'BOE', 'CTO', 'PANDA',
#                'SHARP', 'MDT', 'HKC', 'HSD', 'INX', 'IVO', 'CTC']
#
# TARGET_ICS = ['7272', '7202M', '7202H', '7302']
# YEARS = ['2022年度', '2023年度', '2024年度', '2025年度']
# GRADES = ['v', 'A', 'G', '-', 'NULL']


def statistic_project_status(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', grade='all', publisher='ALL' ,ITO ='ALL'):
    """
    根据条件筛选更新次数的项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :param grade:  /-/v/G/A/ALL
    :return: 项目名，统计结果
    """
    # 数据预处理
    df = df.copy()
    # 处理是否有ITO信息
    site_col = df.iloc[:, 10].astype(str)    #站点
    version_col = df.iloc[:, 14].astype(str)  #版本号信息
    df['Unique_ID'] = df.iloc[:, 0].astype(str) + '_' +df.iloc[:, 2].astype(str) + '_' + df.iloc[:, 3].astype(str) + '_' + df.iloc[:, 5].astype(str) + '_' + df.iloc[:, 7].astype(str)
    # 为每个唯一项目ID标记是否有整机站（Y/N）
    ito_flag = {}  # {Unique_ID: 'Y' or 'N'}
    # 先初始化为 N
    for uid in df['Unique_ID'].unique():
        ito_flag[uid] = 'N'

    # 遍历所有行，发现有效整机站则标记为 Y
    for idx in range(len(df)):
        uid = df.iloc[idx]['Unique_ID']
        site = site_col.iloc[idx]
        version = version_col.iloc[idx]

        # 站点为整机且版本非0（填充后的空值）才视为有效整机站
        site_str = site.strip().lower()
        if site_str in ('整机') and version != '0':
            ito_flag[uid] = 'Y'

    # 将标记添加到df
    df['ITO'] = df['Unique_ID'].map(ito_flag)


    df_filtered = df[df.iloc[:, 10] == '模组'].copy()

    # 标准化各列数据
    def normalize_ic(ic):
        ic = str(ic).upper()
        if '7202M' in ic: return '7202M'
        if '7202H' in ic: return '7202H'
        return ic if ic in TARGET_ICS else 'OTHER'

    def normalize_glass(g):
        match = re.match(r'^([A-Za-z]+)', str(g))
        return match.group(1).upper() if match else str(g).upper()

    # df_filtered.loc[:, 'Sequence'] = df_filtered.iloc[:, 0].astype(str)
    df_filtered.loc[:, 'Grade'] = df_filtered.iloc[:, 1].astype(str).str.upper()
    df_filtered.loc[:, 'IC_Norm'] = df_filtered.iloc[:, 6].apply(normalize_ic)
    df_filtered.loc[:, 'Factory'] = df_filtered.iloc[:, 5].astype(str)
    # df_filtered.loc[:,'Glass_before'] = df.iloc[:, 7].astype(str)
    df_filtered.loc[:, 'Glass_Norm'] = df_filtered.iloc[:, 7].apply(normalize_glass)
    df_filtered.loc[:, 'Flash'] = df_filtered.iloc[:, 9].astype(str).str.upper().apply(
        lambda x: 'Y' if x.startswith('Y') else 'N')
    df_filtered.loc[:, 'Year'] = df_filtered.iloc[:, 16].astype(str).apply(
        lambda x: f"{datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year}年度"
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x)
        else 'UNKNOWN')
    df_filtered.loc[:, 'Half_Year'] = df_filtered.iloc[:, 16].astype(str).apply(
        lambda x: '上半年'
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x)
        and 1 <= datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').month <= 6
        else '下半年'
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x)
        and 7 <= datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').month <= 12
        else 'UNKNOWN')
    df_filtered.loc[:, 'Publisher'] = df_filtered.iloc[:, 17].astype(str)

    df_filtered.loc[:, 'ITO'] = df_filtered['ITO']

    # 创建唯一项目ID
    df_filtered.loc[:, 'Project_ID'] = (
            df_filtered.iloc[:, 0].astype(str) + '_' +
            df_filtered.iloc[:, 2].astype(str) + '_' +
            df_filtered.iloc[:, 3].astype(str) + '_' +
            df_filtered.iloc[:, 5].astype(str) + '_' +
            df_filtered.iloc[:, 7].astype(str)
    )
    # df_filtered['Project_ID'].to_csv('project_ids.txt', index=False, header=False, encoding='utf-8')
    # print("Project_ID 已保存到 project_ids.txt 文件")
    # print(df_filtered.iloc[:, 3].astype(str))


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

    # if year != 'ALL':
    #     # df_filtered = df_filtered[df_filtered['Year'] == year]
    #     valid_projects = df_filtered.drop_duplicates('Project_ID', keep='first')
    #     valid_projects = valid_projects[valid_projects['Year'] == year]['Project_ID']
    #     df_filtered = df_filtered[df_filtered['Project_ID'].isin(valid_projects)]

    if grade != 'ALL':
        # valid_projects = df_filtered.drop_duplicates('Project_ID', keep='first')
        # df_filtered = df_filtered[df_filtered['Grade'] == grade]['Project_ID']
        # df_filtered = df_filtered[df_filtered['Project_ID'].isin(valid_projects)]
        df_filtered = df_filtered[df_filtered['Grade'] == grade]
        # grade_projects = df_filtered[df_filtered['Grade'] == grade]['Project_ID'].unique()
        # df_filtered = df_filtered[df_filtered['Project_ID'].isin(grade_projects)]

    if publisher != 'ALL':
        df_filtered = df_filtered[df_filtered['Publisher'] == publisher]
        # if year != 'ALL':
        #     df_filtered = df_filtered[df_filtered['Year'] == year]
        if year != 'ALL':
            df_filtered = df_filtered[df_filtered['Year'] == year]
        if half_year != 'ALL':
            df_filtered = df_filtered[df_filtered['Half_Year'] == half_year]
    else:
        if year != 'ALL':
            valid_projects = df_filtered.drop_duplicates('Project_ID', keep='first')
            valid_projects = valid_projects[valid_projects['Year'] == year]['Project_ID']
            df_filtered = df_filtered[df_filtered['Project_ID'].isin(valid_projects)]
        if half_year != 'ALL':
            valid_projects = df_filtered.drop_duplicates('Project_ID', keep='first')
            valid_projects = valid_projects[valid_projects['Half_Year'] == half_year]['Project_ID']
            df_filtered = df_filtered[df_filtered['Project_ID'].isin(valid_projects)]

    if ITO != 'ALL':
        df_filtered = df_filtered[df_filtered['ITO'] == ITO]

    # if half_year != 'ALL':
        # df_filtered = df_filtered[df_filtered['Half_Year'] == half_year]

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
        'Year': 'first',
        'Half_Year': 'first',
        'Grade': 'first',
        'Publisher': 'first',
    }).reset_index()

    # 去除更新次数为0的项目统计
    project_stats = project_stats[project_stats['Update_Count'] > 0]
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
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {len(project_stats)}）")
    # print(dist_stats.to_string(index=False))
    print("─" * 40)
    # if (year == '2025年度' and grade == 'G'):
    #     unique_projects = df_filtered['Project_ID'].unique()
    #     for idx, project_id in enumerate(unique_projects, 1):
    #         print(f"{idx}. {project_id}")
    #     print(f"总计: {len(unique_projects)}个项目")
    #     print("─" * 40)

    conditions = []
    if ic_type != 'ALL': conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            conditions.append(f"年度: {year+half_year}")
        else:
            conditions.append(f"年度: {year}")
    # if half_year != 'ALL': conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': conditions.append(f"等级: {grade}")
    if publisher != 'ALL': conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': conditions.append(f"ITO: {ITO}")

    if len(project_stats) == 0:
        print("没有找到符合条件的项目")
        return conditions, pd.Series(dtype=int)  # 返回空Series

    return conditions, dist_stats


def normolization_data(df):
    '''
    规范化excel表MP项目数据结构
    :param df: 二维数组
    :return: 规范后的数据
    '''
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['sequence'] = df.iloc[:, 0].astype(str)
    df['Grade'] = df.iloc[:, 1].astype(str).str.upper()
    df['Project_Name'] = df.iloc[:, 2].astype(str)  # C列
    df['Terminal_Name'] = df.iloc[:, 3].astype(str)  # D列
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
        lambda x: f"{datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').year}年度"
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) else '未知年度'
    )
    df['Half_Year'] = df.iloc[:, 16].astype(str).apply(  # Q列 -> 半年度 H1(上半年)/H2(下半年)
        lambda x: '上半年'
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) 
        and 1 <= datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').month <= 6
        else '下半年'
        if pd.notnull(x) and re.match(r'\d{4}\.\d{1,2}\.\d{1,2}', x) 
        and 7 <= datetime.strptime(x.split(' ')[0].replace('.', '-'), '%Y-%m-%d').month <= 12
        else 'UNKNOWN'
    )
    df['Publisher'] = df.iloc[:, 17].astype(str)
    # 创建唯一项目ID（项目+模组厂+玻璃）
    df['Unique_ID'] = df['sequence'] + '_' + df['Project_Name'] + '_' + df['Terminal_Name'] + '_' + df[
        'Factory'] + '_' + df['Glass_before']
    # df['Unique_ID'].to_csv('project_ids.txt', index=False, header=False, encoding='utf-8')
    # print("Project_ID 已保存到 project_ids.txt 文件")

    # 特殊处理
    target_id = "HW09_图高_海菲_CSOT6.56"
    df.loc[df['Unique_ID'] == target_id, 'Grade'] = 'G'

    # 处理是否有ITO信息
    site_col = df.iloc[:, 10].astype(str)    #站点
    version_col = df.iloc[:, 14].astype(str)  #版本号信息

    # 为每个唯一项目ID标记是否有整机站（Y/N）
    ito_flag = {}  # {Unique_ID: 'Y' or 'N'}
    # 先初始化为 N
    for uid in df['Unique_ID'].unique():
        ito_flag[uid] = 'N'

    # 遍历所有行，发现有效整机站则标记为 Y
    for idx in range(len(df)):
        uid = df.iloc[idx]['Unique_ID']
        site = site_col.iloc[idx]
        version = version_col.iloc[idx]

        # 站点为整机且版本非0（填充后的空值）才视为有效整机站
        site_str = site.strip().lower()
        if site_str in ('整机') and version != '0':
            ito_flag[uid] = 'Y'

    # 将标记添加到df
    df['ITO'] = df['Unique_ID'].map(ito_flag)

    # modified_count = (df['Unique_ID'] == target_id).sum()
    # if modified_count > 0:
    #     print(f"已将 {modified_count} 条记录的Grade修改为G（项目ID: {target_id}）")

    return df


def statistic_ic_projects(df, factory='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', grade='all', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选IC型号项目数量
    :param df: 传入二维数组
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)
    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass.upper()]
    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]
    if year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]
    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

    # IC型号匹配（7202M/7202H模糊匹配）
    # pattern = re.compile(r'7272|7202[MH]|7302')
    pattern = re.compile(r'7202MA|7272CA|7272|7202[MH]|7302|7371|7205C|7274RA')
    df['Matched_IC'] = df['IC_Type'].apply(
        lambda x: pattern.search(x).group() if pattern.search(x) else None
    )

    # 过滤目标IC型号
    filtered_df = df[df['Matched_IC'].isin(TARGET_ICS)]

    # 按唯一ID去重后统计
    ic_counts = filtered_df.drop_duplicates('Unique_ID')['Matched_IC'].value_counts()
    ic_counts = ic_counts.reindex(TARGET_ICS, fill_value=0)
    ic_counts = ic_counts[ic_counts > 0]

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(ic_counts.sum())}）")
    print("─" * 40)

    title_conditions = []
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(ic_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, ic_counts


def statistic_module_factory(df, ic_type='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', grade='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选模组厂项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

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
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(factory_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(factory_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, factory_counts


def statistic_glass_projects(df, ic_type='ALL', factory='ALL', flash='ALL', year='ALL', half_year='ALL', grade='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选玻璃厂项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

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
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(glass_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(glass_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, glass_counts


def statistic_flash_projects(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL', half_year='ALL', grade='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选Flash项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

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
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(flash_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(flash_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series
    return title_conditions, flash_counts


def statistic_project_by_year(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL',  half_year='ALL',grade='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选等级项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    year_counts = unique_projects['Year'].value_counts()
    year_counts = year_counts.reindex(YEARS, fill_value=0)
    # year_counts = year_counts[year_counts > 0]

    # print(year_counts)
    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(year_counts.sum())}）")
    print("─" * 40)
    # if ( grade == 'G'):
    #     unique_projects = df['Unique_ID'].unique()
    #     for idx, project_id in enumerate(unique_projects, 1):
    #         print(f"{idx}. {project_id}")
    #     print(f"总计: {len(unique_projects)}个项目")
    #     print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(year_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, year_counts


def statistic_project_by_half_year(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL', grade='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选半年度（H1上半年/H2下半年）项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param grade:  /-/v/G/A/ALL
    :param publisher: 发布人
    :param ITO: Y/N/ALL
    :return: 标题条件列表，半年度统计结果Series (H1/H2)
    """
    df = normolization_data(df)

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

    if year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计半年度分布
    half_year_counts = unique_projects['Half_Year'].value_counts()
    half_year_counts = half_year_counts.reindex(['上半年', '下半年'], fill_value=0)
    # half_year_counts = half_year_counts[half_year_counts > 0]

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌半年度统计结果（项目总数: {int(half_year_counts.sum())}）")
    # for hy, cnt in half_year_counts.items():
    #     label = half_year_counts
    #     print(f"  {label}: {cnt}个项目")
    # print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    # if year != 'ALL': title_conditions.append(f"年度: {year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(half_year_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, half_year_counts


def statistic_project_by_grade(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', publisher='ALL',ITO = 'ALL'):
    """
    根据条件筛选年度项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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

    if year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    grade_counts = unique_projects['Grade'].value_counts()
    grade_counts = grade_counts.reindex(GRADES, fill_value=0)
    grade_counts = grade_counts[grade_counts > 0]

    # print(year_counts)
    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(grade_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(grade_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, grade_counts


def statistic_project_by_Publisher(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', grade='ALL',ITO = 'ALL'):
    """
    根据条件筛选年度项目数量
    :param df: 传入二维数组
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号，自动提取字母部分并大写
    :param flash: 是否带Flash（Y/N/ALL）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :return: 项目名，统计结果
    """
    df = normolization_data(df)

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

    if year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if ITO != 'ALL':
        df = df[df['ITO'] == ITO]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

    # 按唯一ID去重
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计年度分布
    publisher_counts = unique_projects['Publisher'].value_counts()
    publisher_counts = publisher_counts.reindex(PUBLISHER, fill_value=0)
    publisher_counts = publisher_counts[publisher_counts > 0].sort_values(ascending=False)

    # print(year_counts)
    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"ITO: {ITO if ITO != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(publisher_counts.sum())}）")
    print("─" * 40)

    # 图表美化
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if ITO != 'ALL': title_conditions.append(f"ITO: {ITO}")

    if len(publisher_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, publisher_counts



def statistic_ito_projects(df, ic_type='ALL', factory='ALL', glass='ALL', flash='ALL', year='ALL', half_year='ALL', grade='ALL', publisher='ALL'):
    """
    统计项目是否有整机站（ITO）
    :param df: 原始数据DataFrame
    :param ic_type: 7272/7202M/7202H/ALL
    :param factory: 模组厂
    :param glass: 玻璃型号（自动提取字母部分并大写）
    :param year: 2023/2024/2025/ALL
    :param half_year: 半年度 H1(上半年)/H2(下半年)/ALL(全年)
    :param grade: 等级
    :param publisher: 发布人
    :return: (标题条件列表, 统计结果Series) 结果包含 Y（有整机站）和 N（无整机站）的计数
    """
    # 标准化数据
    df = normolization_data(df)

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
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Year'] == year]

    if grade != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Grade'] == grade]

    if publisher != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Publisher'] == publisher]

    if flash != 'ALL':
        df = df[df['Flash'] == flash.upper()]

    if half_year != 'ALL':
        df = df.drop_duplicates('Unique_ID', keep='first')
        df = df[df['Half_Year'] == half_year]

    # 按唯一ID去重，保留每个项目的第一条记录
    unique_projects = df.drop_duplicates('Unique_ID')

    # 统计ITO标志
    ito_counts = unique_projects['ITO'].value_counts()
    ito_counts = ito_counts.reindex(['Y', 'N'], fill_value=0)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"是否带Flash: {flash if flash != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"半年度: {half_year if half_year != 'ALL' else '全部'}")
    print(f"等级: {grade if grade != 'ALL' else '全部'}")
    print(f"发布人: {publisher if publisher != 'ALL' else '全部'}")
    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(ito_counts.sum())}）")
    print("─" * 40)

    # 构建标题条件
    title_conditions = []
    if ic_type != 'ALL': title_conditions.append(f"IC: {ic_type}")
    if factory != 'ALL': title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL': title_conditions.append(f"玻璃: {glass}")
    if flash != 'ALL': title_conditions.append(f"Flash: {flash}")
    if year != 'ALL':
        if half_year != 'ALL':
            title_conditions.append(f"年度: {year+half_year}")
        else:
            title_conditions.append(f"年度: {year}")
    # if half_year != 'ALL': title_conditions.append(f"半年度: {half_year}")
    if grade != 'ALL': title_conditions.append(f"等级: {grade}")
    if publisher != 'ALL': title_conditions.append(f"发布人: {publisher}")

    if len(ito_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, ito_counts