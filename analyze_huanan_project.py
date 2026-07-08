#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：MP TOOL PROJECT_v5.4_20260122 
@File    ：analyze_huanan_project.py
@Author  ：zxh
@Date    ：2026/1/22 16:44 
'''

from read_excel import *

data_file = 'compare_information.xlsx'
datas_HN = read_compare_excel(data_file, sheet_names='Sheet3')
IC_TYPE = datas_HN.iloc[:, 0].dropna().astype(str).tolist()
FACTORIES = datas_HN.iloc[:, 1].dropna().tolist()
MONTHS_PROJECT = datas_HN.iloc[:, 2].dropna().tolist()
# print(MONTHS_PROJECT)
GLASS_TYPE = datas_HN.iloc[:, 3].dropna().tolist()
TP_MODE = datas_HN.iloc[:, 4].dropna().tolist()
TP_MODE.append('NULL')
PALM_TYPE = datas_HN.iloc[:, 5].dropna().tolist()
PALM_TYPE.append('NULL')
STATUS_PROJECT = datas_HN.iloc[:, 6].dropna().tolist()
CLASS_PROJECT = datas_HN.iloc[:, 7].dropna().tolist()
CLASS_PROJECT.append('NULL')
# print(CLASS_PROJECT)
SCALE_NAME = datas_HN.iloc[:, 8].dropna().tolist()
PM_MANE = datas_HN.iloc[:, 9].dropna().tolist()
YEAR_PROJECT = datas_HN.iloc[:, 10].dropna().tolist()


def parse_mixed_date_column_to_year_month(analyse_type, data, column, date_format='%Y年%m月%d日',
                                          excel_origin='1899-12-30'):
    """
    解析混合格式的日期列，根据analyse_type返回年份或月份
    :param analyse_type: '年度'返回年度，'月度'返回年月，'日期'返回完整日期
    :param data: 原始数据(DataFrame或Series)
    :param column: 对应列(列名或列索引)
    :param date_format: 文本日期格式
    :param excel_origin: excel表公式日期基准
    :return: 解析后的日期Series和格式化后的Series
    """
    # 提取目标列数据
    if isinstance(data, pd.DataFrame):
        if isinstance(column, int):
            date_series = data.iloc[:, column]
        else:
            date_series = data[column]
    elif isinstance(data, pd.Series):
        date_series = data
    else:
        raise ValueError("data参数必须是DataFrame或Series")

    # 处理None或NaN值
    date_series = date_series.copy()

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

    # 根据analyse_type返回不同格式
    if analyse_type == '年度':
        # 返回年份
        formatted_result = parsed_dates.dt.year
        formatted_result = formatted_result.apply(lambda x: f"{int(x)}年度" if pd.notna(x) else None)

    elif analyse_type == '月度':
        # 返回年月（格式：YYYY年MM月）
        formatted_result = parsed_dates.apply(
            lambda x: f"{x.month:d}月" if pd.notna(x) else None
        )

    elif analyse_type == '日期':
        # 返回完整日期（格式：YYYY年MM月DD日）
        formatted_result = parsed_dates.apply(
            lambda x: f"{x.month:02d}月{x.day:02d}日" if pd.notna(x) else None
        )

    elif analyse_type == '季度':
        # 返回季度（格式：YYYY年Qn季度）
        formatted_result = parsed_dates.apply(
            lambda x: f"{x.year}年Q{(x.month - 1) // 3 + 1}季度" if pd.notna(x) else None
        )

    elif analyse_type == '年月':
        # 返回年月（格式：YYYY-MM）
        formatted_result = parsed_dates.apply(
            lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else None
        )

    else:
        raise ValueError(f"不支持的analyse_type: {analyse_type}，请使用'年度'、'月度'、'日期'、'季度'或'年月'")

    return parsed_dates, formatted_result


def normolization_dataHN(df):
    '''
    规范化excel表客诉记录数据结构
    :param df: 二维数组
    :return: 规范后的数据
    '''
    # 数据预处理
    df = df.copy()

    # 标准化各列数据
    df['sequence'] = df.iloc[:, 0].astype(str)  # 序号
    df['IC_Type'] = df.iloc[:, 1].astype(str)  # IC
    df['Factory'] = df.iloc[:, 3].astype(str)  # 模组厂
    df['Glass'] = df.iloc[:, 5].astype(str)  # 玻璃
    parsed_dates, formatted_month = parse_mixed_date_column_to_year_month('月度', df, 4)
    df['Month'] = formatted_month.astype(str)  # 月份
    parsed_dates, formatted_year = parse_mixed_date_column_to_year_month('年度', df, 4)
    df['Year'] = formatted_year.astype(str)  # 年份
    df['TP_Mode'] = df.iloc[:, 8].astype(str)  # TP mode
    df['Palm'] = df.iloc[:, 11].astype(str)  # 开关贴脸息屏
    df['Status'] = df.iloc[:, 13].astype(str)  # 项目状态
    df['Category'] = df.iloc[:, 14].astype(str)  # 类别
    df['Sales'] = df.iloc[:, 15].astype(str)  # 销售
    df['PM'] = df.iloc[:, 16].astype(str)  # PM

    return df


def statistic_ic_projects_HN(df, factory='ALL', glass='ALL', year='ALL', month='ALL', TP_mode='ALL', palm='ALL', status='ALL', category='ALL',sales='ALL', pm='ALL'):
    """
    以IC为X轴统计项目数量
    :param df: 原始数据DataFrame
    :param factory: 模组厂筛选条件，默认'ALL'表示全部
    :param glass: 玻璃型号筛选条件，默认'ALL'表示全部
    :param year: 年度筛选条件，默认'ALL'表示全部
    :param month: 月份筛选条件，默认'ALL'表示全部
    :param TP_mode: TP模式筛选条件，默认'ALL'表示全部
    :param palm: palm筛选条件，默认'ALL'表示全部
    :param status: 项目状态筛选条件，默认'ALL'表示全部
    :param category: 项目类别筛选条件，默认'ALL'表示全部
    :param sales: 销售筛选条件，默认'ALL'表示全部
    :param pm: PM筛选条件，默认'ALL'表示全部
    :return: 筛选条件列表和IC统计结果
    """
    # 标准化各列数据
    df = normolization_dataHN(df)

    # 应用筛选条件
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # 只保留定义的IC型号
    # df = df[df['IC_Type'].isin(IC_TYPE)]
    df['IC_Type'] = df['IC_Type'].apply(lambda x:x if x in IC_TYPE else 'Unkonwn')
    # 统计IC
    IC_counts = df['IC_Type'].value_counts()
    all_ic_types = IC_TYPE+['Unkonwn']
    IC_counts = IC_counts.reindex(all_ic_types, fill_value=0)
    IC_counts = IC_counts[IC_counts > 0].sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(IC_counts.sum())}）")
    print("─" * 40)

    IC_counts_sum =  int(IC_counts.sum())

    title_conditions = []
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(IC_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)  # 返回空Series

    return title_conditions, IC_counts, IC_counts_sum


def statistic_factory_projects_HN(df, ic_type='ALL', glass='ALL', year='ALL', month='ALL',
                                  TP_mode='ALL', palm='ALL', status='ALL', category='ALL',
                                  sales='ALL', pm='ALL'):
    """
    以模组厂为X轴统计项目数量
    :param df: 原始数据DataFrame
    :param ic_type: IC型号筛选条件，默认'ALL'表示全部
    :param glass: 玻璃型号筛选条件，默认'ALL'表示全部
    :param year: 年度筛选条件，默认'ALL'表示全部
    :param month: 月份筛选条件，默认'ALL'表示全部
    :param TP_mode: TP模式筛选条件，默认'ALL'表示全部
    :param palm: 手掌类型筛选条件，默认'ALL'表示全部
    :param status: 项目状态筛选条件，默认'ALL'表示全部
    :param category: 项目类别筛选条件，默认'ALL'表示全部
    :param sales: 销售渠道筛选条件，默认'ALL'表示全部
    :param pm: PM负责人筛选条件，默认'ALL'表示全部
    :return: 筛选条件列表和模组厂统计结果
    """
    # 标准化各列数据
    df = normolization_dataHN(df)

    # 应用筛选条件
    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # # 找出不在FACTORIES列表中的模组厂字符
    # all_factories = df['Factory'].dropna().unique()
    # not_in_factories = [factory for factory in all_factories if factory not in FACTORIES]
    #
    # # 打印不在FACTORIES列表中的字符
    # if not_in_factories:
    #     print("▌ 以下模组厂不在预定义列表中:")
    #     for factory in not_in_factories:
    #         print(f"  - '{factory}'")
    #     print("─" * 40)

    # 只保留定义的模组厂
    # df = df[df['Factory'].isin(FACTORIES)]

    # 将不在FACTORIES中的值替换为'Unknown'
    df['Factory'] = df['Factory'].apply(lambda x: x if x in FACTORIES else 'Unknown')
    # 统计模组厂
    factory_counts = df['Factory'].value_counts()
    all_factories = FACTORIES + ['Unknown']
    factory_counts = factory_counts.reindex(all_factories, fill_value=0)
    factory_counts = factory_counts[factory_counts > 0].sort_values(ascending=False)

    # 打印筛选条件
    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(factory_counts.sum())}）")
    print("─" * 40)

    factory_counts_sum = int(factory_counts.sum())

    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(factory_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, factory_counts, factory_counts_sum


def statistic_glass_projects_HN(df, ic_type='ALL', factory='ALL', year='ALL', month='ALL',
                                TP_mode='ALL', palm='ALL', status='ALL', category='ALL',
                                sales='ALL', pm='ALL'):
    """
    以玻璃型号为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # df = df[df['Glass'].isin(GLASS_TYPE)]

    df['Glass'] = df['Glass'].apply(lambda x: x if x in GLASS_TYPE else 'Unknown')
    # 统计玻璃型号
    glass_counts = df['Glass'].value_counts()
    all_glass_types = GLASS_TYPE + ['Unknown']
    glass_counts = glass_counts.reindex(all_glass_types, fill_value=0)
    glass_counts = glass_counts[glass_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(glass_counts.sum())}）")
    print("─" * 40)

    glass_counts_sum = int(glass_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(glass_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, glass_counts, glass_counts_sum


def statistic_year_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', month='ALL',
                               TP_mode='ALL', palm='ALL', status='ALL', category='ALL',
                               sales='ALL', pm='ALL'):
    """
    以年度为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    df = df[df['Year'].isin(YEAR_PROJECT)]
    year_counts = df['Year'].value_counts()
    year_counts = year_counts.reindex(YEAR_PROJECT, fill_value=0)
    year_counts = year_counts[year_counts > 0]

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(year_counts.sum())}）")
    print("─" * 40)

    year_counts_sum = int(year_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(year_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, year_counts, year_counts_sum


def statistic_month_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                                TP_mode='ALL', palm='ALL', status='ALL', category='ALL',
                                sales='ALL', pm='ALL'):
    """
    以月份为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    df = df[df['Month'].isin(MONTHS_PROJECT)]
    month_counts = df['Month'].value_counts()
    month_counts = month_counts.reindex(MONTHS_PROJECT, fill_value=0)
    month_counts = month_counts[month_counts > 0]

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(month_counts.sum())}）")
    print("─" * 40)

    month_counts_sum = int(month_counts.sum())

    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(month_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, month_counts, month_counts_sum


def statistic_tpmode_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                                 month='ALL', palm='ALL', status='ALL', category='ALL',
                                 sales='ALL', pm='ALL'):
    """
    以TP模式为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # df = df[df['TP_Mode'].isin(TP_MODE)]
    df['TP_Mode'] = df['TP_Mode'].apply(lambda x: x if x in TP_MODE else 'Unknown')
    # 统计TP模式
    tpmode_counts = df['TP_Mode'].value_counts()
    all_tpmodes = TP_MODE + ['Unknown']
    tpmode_counts = tpmode_counts.reindex(all_tpmodes, fill_value=0)
    tpmode_counts = tpmode_counts[tpmode_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(tpmode_counts.sum())}）")
    print("─" * 40)

    tpmode_counts_sum = int(tpmode_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(tpmode_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, tpmode_counts,tpmode_counts_sum


def statistic_palm_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                               month='ALL', TP_mode='ALL', status='ALL', category='ALL',
                               sales='ALL', pm='ALL'):
    """
    以palm为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # df = df[df['Palm'].isin(PALM_TYPE)]
    df['Palm'] = df['Palm'].apply(lambda x: x if x in PALM_TYPE else 'Unknown')
    # 统计palm
    palm_counts = df['Palm'].value_counts()
    all_palms = PALM_TYPE + ['Unknown']
    palm_counts = palm_counts.reindex(all_palms, fill_value=0)
    palm_counts = palm_counts[palm_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(palm_counts.sum())}）")
    print("─" * 40)

    palm_counts_sum = int(palm_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(palm_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, palm_counts, palm_counts_sum


def statistic_status_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                                 month='ALL', TP_mode='ALL', palm='ALL', category='ALL',
                                 sales='ALL', pm='ALL'):
    """
    以项目状态为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # df = df[df['Status'].isin(STATUS_PROJECT)]
    df['Status'] = df['Status'].apply(lambda x: x if x in STATUS_PROJECT else 'Unknown')
    status_counts = df['Status'].value_counts()
    all_status = STATUS_PROJECT + ['Unknown']
    status_counts = status_counts.reindex(all_status, fill_value=0)
    status_counts = status_counts[status_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(status_counts.sum())}）")
    print("─" * 40)

    status_counts_sum = int(status_counts.sum())

    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(status_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, status_counts, status_counts_sum


def statistic_category_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                                   month='ALL', TP_mode='ALL', palm='ALL', status='ALL',
                                   sales='ALL', pm='ALL'):
    """
    以项目类别为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    # df = df[df['Category'].isin(CLASS_PROJECT)]
    df['Category'] = df['Category'].apply(lambda x: x if x in CLASS_PROJECT else 'Unknown')
    category_counts = df['Category'].value_counts()
    all_categories = CLASS_PROJECT + ['Unknown']
    category_counts = category_counts.reindex(all_categories, fill_value=0)
    category_counts = category_counts[category_counts > 0]

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(category_counts.sum())}）")
    print("─" * 40)

    category_counts_sum = int(category_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(category_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, category_counts,category_counts_sum


def statistic_sales_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                                month='ALL', TP_mode='ALL', palm='ALL', status='ALL',
                                category='ALL', pm='ALL'):
    """
    以销售为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if pm != 'ALL':
        df = df[df['PM'] == pm]

    df = df[df['Sales'].isin(SCALE_NAME)]
    sales_counts = df['Sales'].value_counts()
    sales_counts = sales_counts.reindex(SCALE_NAME, fill_value=0)
    sales_counts = sales_counts[sales_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"PM负责人: {pm if pm != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(sales_counts.sum())}）")
    print("─" * 40)

    sales_counts_sum = int(sales_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if pm != 'ALL':
        title_conditions.append(f"PM负责人: {pm}")

    if len(sales_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, sales_counts, sales_counts_sum


def statistic_pm_projects_HN(df, ic_type='ALL', factory='ALL', glass='ALL', year='ALL',
                             month='ALL', TP_mode='ALL', palm='ALL', status='ALL',
                             category='ALL', sales='ALL'):
    """
    以PM为X轴统计项目数量
    """
    df = normolization_dataHN(df)

    if ic_type != 'ALL':
        df = df[df['IC_Type'] == ic_type]
    if factory != 'ALL':
        df = df[df['Factory'] == factory]
    if glass != 'ALL':
        df = df[df['Glass'] == glass]
    if year != 'ALL':
        df = df[df['Year'] == year]
    if month != 'ALL':
        df = df[df['Month'] == month]
    if TP_mode != 'ALL':
        df = df[df['TP_Mode'] == TP_mode]
    if palm != 'ALL':
        df = df[df['Palm'] == palm]
    if status != 'ALL':
        df = df[df['Status'] == status]
    if category != 'ALL':
        df = df[df['Category'] == category]
    if sales != 'ALL':
        df = df[df['Sales'] == sales]

    df = df[df['PM'].isin(PM_MANE)]
    pm_counts = df['PM'].value_counts()
    pm_counts = pm_counts.reindex(PM_MANE, fill_value=0)
    pm_counts = pm_counts[pm_counts > 0].sort_values(ascending=False)

    print("▌ 筛选条件：")
    print(f"IC型号: {ic_type if ic_type != 'ALL' else '全部'}")
    print(f"模组厂: {factory if factory != 'ALL' else '全部'}")
    print(f"玻璃型号: {glass if glass != 'ALL' else '全部'}")
    print(f"年度: {year if year != 'ALL' else '全部'}")
    print(f"月份: {month if month != 'ALL' else '全部'}")
    print(f"TP模式: {TP_mode if TP_mode != 'ALL' else '全部'}")
    print(f"手掌类型: {palm if palm != 'ALL' else '全部'}")
    print(f"项目状态: {status if status != 'ALL' else '全部'}")
    print(f"项目类别: {category if category != 'ALL' else '全部'}")
    print(f"销售渠道: {sales if sales != 'ALL' else '全部'}")

    print("─" * 40)
    print(f"▌统计结果（项目总数: {int(pm_counts.sum())}）")
    print("─" * 40)

    pm_counts_sum =int(pm_counts.sum())
    title_conditions = []
    if ic_type != 'ALL':
        title_conditions.append(f"IC型号: {ic_type}")
    if factory != 'ALL':
        title_conditions.append(f"模组厂: {factory}")
    if glass != 'ALL':
        title_conditions.append(f"玻璃: {glass}")
    if year != 'ALL':
        title_conditions.append(f"年度: {year}")
    if month != 'ALL':
        title_conditions.append(f"月份: {month}")
    if TP_mode != 'ALL':
        title_conditions.append(f"TP模式: {TP_mode}")
    if palm != 'ALL':
        title_conditions.append(f"手掌类型: {palm}")
    if status != 'ALL':
        title_conditions.append(f"项目状态: {status}")
    if category != 'ALL':
        title_conditions.append(f"项目类别: {category}")
    if sales != 'ALL':
        title_conditions.append(f"销售渠道: {sales}")

    if len(pm_counts) == 0:
        print("没有找到符合条件的项目")
        return title_conditions, pd.Series(dtype=int)

    return title_conditions, pm_counts, pm_counts_sum



# debug
# file_name = 'LCD TDDI华南项目汇总_0122.xlsx'
# df = read_excel_for_analyseHN(file_name)


# 合并获取第5列数据并去重
# df_date = df.iloc[:,4]

# parsed_dates, date_formatted = parse_mixed_date_column_to_year_month('年度', df,4)
# print(date_formatted)
# date_formatted.drop_duplicates(inplace=True)   # 去除所有重复行
# date_formatted.to_csv('year.txt', index=False, header=False, encoding='utf-8')
