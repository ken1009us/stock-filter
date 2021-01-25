# stock filter by Ken Wu
import requests as rq
import pandas as pd
import xlsxwriter
import matplotlib.pyplot as plt
import os
import shutil

from bs4 import BeautifulSoup

# Make the folder to save your reports
path = './stock_filter_pdf_file'
if os.path.isdir(path):
    shutil.rmtree('./stock_filter_pdf_file')
    os.mkdir('./stock_filter_pdf_file')
else:
    os.mkdir('./stock_filter_pdf_file')

# Make user input the info
holder = input("Please enter the number of stocks that the shareholders hold (1 = 4 lots, 2 = 10 lots): ")
week = input("Please enter the growth week (>= 0~12 week): ")
before_week = input("Compare with n weeks ago (0~11 week): ")
growth_rate = input("Please enter the growth rate(%) <=-3, <=-2, <=-1, <=-0.01, >=0.1, >=1, >=2, >=3: ")
stock_price = input("Please enter the current stock price <= 20, 30, 40, 50 ,60 ,70, 80, 90, 100, 1000, 5000: ")

# crawler
url = f'https://norway.twsthr.info/StockHoldersContinue.aspx?Show={holder}&continue=Y&weeks={week}&growthrate={growth_rate}&beforeweek={before_week}&price={stock_price}&valuerank=1-3000&display=0'
res = rq.get(url)
res.encoding = 'utf-8'
soup = BeautifulSoup(res.text, 'html.parser')
rows = soup.find_all('tr')
info_list = []
for row in rows: 
    stock_list = []
    for rr in row: 
        stock_list.append(rr.string) 
    info_list.append(stock_list[3:6])
f_list = info_list[18:-1]

s_list = []
for final_stock in f_list:
    if '櫃' not in final_stock[1]:
        s_list.append(final_stock)

stock_num_name = []
stock_name = []
stock_num = []
stock_focus_url = []
stock_holder_url = []
category = []
percent = []
for spe_info in s_list:
    stock_num_name.append(spe_info[0])
    category.append(spe_info[1])
    percent.append(spe_info[2])
for n in stock_num_name:
    num = n.split(' ')[0]
    name = n.split(' ')[1]
    stock_num.append(num)
    stock_name.append(name)
for fh in stock_num:
    f_url = f'https://histock.tw/stock/main.aspx?no={fh}'
    stock_focus_url.append(f_url)
    h_url = f'https://norway.twsthr.info/StockHolders.aspx?STOCK={fh}'
    stock_holder_url.append(h_url)

# to excel
def data_to_df(stock_num, stock_name, category, percent, stock_focus_url):
    df = pd.DataFrame({'number': stock_num, 
                       'stock': stock_name, 
                       'category': category, 
                       'percent': percent,
                       'url_1': stock_focus_url
                      #'url_2': stock_holder_url
                        })
    # print(df)
    df.to_excel('main_report.xlsx', engine='xlsxwriter', index=False) 

# revenue reports
def financial_report(stock_num):
    df_list = []
    # n = 0
    re_name = 0
    for num in stock_num:
        fin_table_list = []
        financial_url = f'https://histock.tw/stock/{num}/%E8%B2%A1%E5%8B%99%E5%A0%B1%E8%A1%A8'
        fin_res = rq.get(financial_url)
        fin_res.encoding = 'utf-8'
        fin_soup = BeautifulSoup(fin_res.text, 'html.parser')
        fin_rows = fin_soup.find_all('tr')
        for fin_row in fin_rows: 
            fin_table = []
            for sin_fin_row in fin_row: 
                fin_table.append(sin_fin_row.string)
            fin_table_list.append(fin_table)

        fin_table_list = fin_table_list[2:]
        month = []
        sin_mon_re = []
        last_year_mon_re = []
        mom = []
        yoy = []
        for content in fin_table_list:
            month.append(content[0])
            sin_mon_re.append(content[1])
            last_year_mon_re.append(content[2])
            mom.append(float(content[3].split('%')[0]))
            yoy.append(float(content[4].split('%')[0]))

        re_df = pd.DataFrame({'month': month, 
                           'month revenue': sin_mon_re, 
                           'same month last year': last_year_mon_re, 
                           'MoM': mom,
                           'YoY': yoy})
        dec_re_df = re_df.sort_index(ascending=False)
        df_list.append(dec_re_df)
        # print(df_list)

        ff, aa = plt.subplots(figsize=(10,5))
        aa.set_xticklabels('month', rotation=45, ha='center', fontsize=6)
        ab = aa.twinx()
        aa.set_ylabel('MoM',color='black')
        curve1, = aa.plot(dec_re_df.month, dec_re_df.MoM, 'o-', 
                          label="MoM", 
                          color='Blue')

        ab.set_ylabel('YoY', color='black')
        curve2, = ab.plot(dec_re_df.month, dec_re_df.YoY, 'o-', 
                          label="YoY", 
                          color='red')
        figure = [curve1, curve2]
        aa.legend(figure, [ima.get_label() for ima in figure], bbox_to_anchor=(1.1, 1), loc=2, fontsize=8)
        plt.title("Revenue Filter by Ken Wu" + '(TW ' + stock_num[re_name] + ')', 
                  fontdict={'family': 'serif', 
                            'color' : 'darkblue',
                            'weight': 'bold',
                            'size': 14})
        aa.grid(True)
        ff.subplots_adjust(right=0.7)
        # plt.show()
        ff.savefig('./stock_filter_pdf_file/%s_%s_revenue.pdf' % (stock_num[re_name], stock_name[re_name]))
        re_name += 1

# Merge to pdf file and make the bar graph/line graph
def stock_holder_filter_to_pdf(stock_num, stock_name, stock_holder_url):
    file_name = 0
    for tri_url in stock_holder_url:
        table_list = []
        date_list = []
        total_holder_list = []
        more_than_four_hundred_list = []
        tri_res = rq.get(tri_url)
        tri_res.encoding = 'utf-8'
        tri_soup = BeautifulSoup(tri_res.text, 'html.parser')
        tri_rows = tri_soup.find_all('tr')
        for tri_row in tri_rows: 
            holder_table = []
            for sin_row in tri_row: 
                holder_table.append(sin_row.string)
            table_list.append(holder_table)
        for elements in table_list[16:64]:
            # [16:356]:three years
            # [16:123]:one year
            # [16:64]:half year
            if "\xa0" in elements:
                date = elements[2].split("\xa0")[0]
                date_list.append(date)
                total_holder = float(elements[4].split("\\")[0].replace(',', ''))
                total_holder_list.append(total_holder)
                more_than_four_hundred = float(elements[7])
                more_than_four_hundred_list.append(more_than_four_hundred)
        tri_df = pd.DataFrame({'date': date_list,
                               'total_shareholders': total_holder_list,
                               'more_than_400_shareholders': more_than_four_hundred_list})
        dec_df = tri_df.sort_index(ascending=False)

        fig, ax1 = plt.subplots()
        ax1.set_xticklabels('date', rotation=45, ha='center', fontsize=6)
        ax2 = ax1.twinx()
        ax1.set_ylabel('total shareholders',color='black')
        bar = ax1.bar(dec_df.date, dec_df.total_shareholders, 
                      label="total shareholders", 
                      edgecolor='black', 
                      color='skyblue')
        ax2.set_ylabel('>400 shareholders', color='black')
        curve, = ax2.plot(dec_df.date, dec_df.more_than_400_shareholders, 'o-', 
                          label=">400 shareholders", 
                          color='tomato')
        figure = [bar, curve]
        ax1.legend(figure, [ima.get_label() for ima in figure], bbox_to_anchor=(1.1, 1), loc=2, fontsize=8)
        plt.title("Stock Filter by Ken Wu" + '(TW ' + stock_num[file_name] + ')', 
                  fontdict={'family': 'serif', 
                            'color' : 'darkblue',
                            'weight': 'bold',
                            'size': 14})
        ax1.grid(True)
        fig.subplots_adjust(right=0.7)
        # plt.show()
        fig.savefig('./stock_filter_pdf_file/%s_%s_holder.pdf' % (stock_num[file_name], stock_name[file_name]))
        file_name += 1

def main():
    data_to_df(stock_num, stock_name, category, percent, stock_focus_url)
    stock_holder_filter_to_pdf(stock_num, stock_name, stock_holder_url)
    financial_report(stock_num)

main()