# -*- coding: utf8 -*-
# author: Zhou Mengying
# 2017-12-27
import seaborn as sns
from WindPy import w
import matplotlib.pyplot as plt
import pandas as pd
import xlrd

plt.rcParams['font.sans-serif']=['SimHei'] # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False # 用来正常显示负号


def draw_graph(allocation_result, selection_result, interaction_result):
    # 画结果图
    sns.set_color_codes("muted")
    # draw allocation
    pd_df = pd.DataFrame(list(allocation_result.items()))
    pd_df.columns = ["industry", "allocation"]
    pd_df = pd_df.sort_values(["allocation"], ascending=False).reset_index(drop=True)
    plt.figure(figsize=(16/3, 3))
    ax = sns.barplot(pd_df.allocation, pd_df.industry, color="b")
    ax.set(xlabel="allocation", ylabel="industry")
    ax.get_xaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(round(x, 6))))
    ax.set_yticklabels(pd_df.industry)
    ax.set(ylabel="", xlabel="行业配置收益")
    plt.tight_layout()
    plt.savefig("brinson_model_allocation.png", dpi=120)

    # draw selection
    pd_df = pd.DataFrame(list(selection_result.items()))
    pd_df.columns = ["industry", "selection"]
    pd_df = pd_df.sort_values(["selection"], ascending=False).reset_index(drop=True)
    plt.figure(figsize=(16/3, 3))
    ax = sns.barplot(pd_df.selection, pd_df.industry, color="y")
    ax.set(xlabel="selection", ylabel="industry")
    ax.get_xaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(round(x, 6))))
    ax.set_yticklabels(pd_df.industry)
    ax.set(ylabel="", xlabel="个股选择收益")
    plt.tight_layout()
    plt.savefig("brinson_model_selection.png", dpi=120)

    # draw interaction
    pd_df = pd.DataFrame(list(interaction_result.items()))
    pd_df.columns = ["industry", "interaction"]
    pd_df = pd_df.sort_values(["interaction"], ascending=False).reset_index(drop=True)
    plt.figure(figsize=(16/3, 3))
    ax = sns.barplot(pd_df.interaction, pd_df.industry, color="r")
    ax.set(xlabel="interaction", ylabel="industry")
    ax.get_xaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(round(x, 6))))
    ax.set_yticklabels(pd_df.industry)
    ax.set(ylabel="", xlabel="综合收益")
    plt.tight_layout()
    plt.savefig("brinson_model_interation.png", dpi=120)


if __name__ == '__main__':
    data = xlrd.open_workbook('SL4856_华觉多策略二号私募证券投资基金_资产估值表_20171208.xls')
    sheet = data.sheets()[0] # get the first sheet dat
    nrows = sheet.nrows
    ncols = sheet.ncols

    stock_list = []
    w.start()
    for i in range(nrows):
        row_data = sheet.row_values(i) # row_data[13] is 估值增值 and row_data[0] is 科目代码
        if row_data[13] != '' and (row_data[0].endswith('SH') or row_data[0].endswith('SZ')):
            temp_list = []
            stock_code = '{0}.{1}'.format(row_data[0][-9:-3], row_data[0][-2:]) # reformat to Wind stock code
            stock_name = row_data[1]
            wsd_data = w.wsd(stock_code, "industry_sw", "2017-12-08", "2017-12-08", "industryType=1")
            stock_industry = wsd_data.Data[0][0]
            cost = row_data[7]
            weight = round(row_data[8], 6)
            fund_yield_rate = round(row_data[14] / cost, 6)

            temp_list.append(stock_code) # 0
            temp_list.append(stock_name) # 1
            temp_list.append(stock_industry) # 2
            temp_list.append(cost) # 3
            temp_list.append(weight) # 4
            temp_list.append(fund_yield_rate) # 5

            stock_list.append(temp_list)

    fund_weight_dic = {}
    for ele in stock_list:
        industry_key = ele[2]
        if fund_weight_dic.get(industry_key) is None:
            fund_weight_dic[industry_key] = float(ele[4])
        else:
            fund_weight_dic[industry_key] = float(fund_weight_dic.get(industry_key)) + float(ele[4])

    fund_yield_dic = {}
    for ele in stock_list:
        industry_key = ele[2]
        if fund_yield_dic.get(industry_key) is None:
            fund_yield_dic[industry_key] = float(ele[5])
        else:
            fund_yield_dic[industry_key] = float(fund_yield_dic.get(industry_key)) + float(ele[5])

    # get benchmark info: "date=2017-12-08;windcode=000300.SH"
    sheet = data.sheets()[1] # get the second sheet data, ensure you add benchmark data as second sheet
    nrows = sheet.nrows
    ncols = sheet.ncols

    bench_weight_dic = {}
    for i in range(nrows):
        row_data = sheet.row_values(i)
        industry_key = row_data[0]
        if bench_weight_dic.get(industry_key) is None:
            bench_weight_dic[industry_key] = float(row_data[2])/100.0
        else:
            bench_weight_dic[industry_key] = bench_weight_dic.get(industry_key) + float(row_data[2])/100.0

    bench_yield_dic = {}
    for i in range(nrows):
        row_data = sheet.row_values(i)
        industry_key = row_data[0]
        if bench_yield_dic.get(industry_key) is None:
            bench_yield_dic[industry_key] = float(row_data[3])/100.0
        else:
            bench_yield_dic[industry_key] = bench_yield_dic.get(industry_key) + float(row_data[3])/100.0

    allocation_result = {}
    for industry_key in fund_yield_dic.keys():
        allocation_result[industry_key] = round((fund_weight_dic.get(industry_key) - bench_weight_dic.get(industry_key)) \
                                          * bench_yield_dic.get(industry_key), 6)

    selection_result = {}
    for industry_key in fund_yield_dic.keys():
        selection_result[industry_key] = round((fund_yield_dic[industry_key] - bench_yield_dic[industry_key]) \
                                         * bench_weight_dic[industry_key], 6)

    interaction_result = {}
    for industry_key in fund_yield_dic.keys():
        interaction_result[industry_key] = round(fund_yield_dic[industry_key] * fund_weight_dic[industry_key] \
                                           - bench_yield_dic[industry_key] * bench_weight_dic[industry_key], 6)

    draw_graph(allocation_result, selection_result, interaction_result)
