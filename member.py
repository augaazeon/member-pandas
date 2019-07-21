import pandas as pd
import numpy as np
from openpyxl import load_workbook
SalesDetails = pd.read_excel(r'F:\会员店0718\恒兴7.1-7.14销售明细.xlsx',dtype=str)
SalesDetails['实收金额'] = SalesDetails['实收金额'].str.replace(',', '').astype('float64')#浮点型#货币格式转换
SalesDetails['开单日期'] = pd.to_datetime(SalesDetails['开单日期'])
SalesDetails['数量'] = SalesDetails['数量'].astype(int)

#销售数量
SalesVolume = SalesDetails[['商品代码','数量']].groupby(by=['商品代码'],as_index=False).sum()#  数据类型问题
#动销客户数
Number_of_customers = SalesDetails[['客户代码','商品代码']].drop_duplicates(['客户代码','商品代码'],keep='first').groupby(by=['商品代码'],as_index=False).count()
#动销频次
Frequency_of_sales = SalesDetails[['开单日期','商品代码']].drop_duplicates(['开单日期','商品代码'],keep='first').groupby(by=['商品代码'],as_index=False).count()

df1 = pd.merge(SalesVolume,Number_of_customers,on='商品代码',how='left')
df2 = pd.merge(df1,Frequency_of_sales,on='商品代码',how='left')  #过渡表，提供给‘战略品种’



deadline = pd.to_datetime('2019-07-07')#截止日期
deadline1 = pd.to_datetime('2019-07-14')
deadline2 = pd.to_datetime('2019-07-21')
deadline3 = pd.to_datetime('2019-07-28')
deadline4 = pd.to_datetime('2019-07-31')
#根据截止日期，以客户为主键，字段：销售额、品规数
def sale_product(deadline):
    Sales = SalesDetails[SalesDetails['开单日期']<=deadline][['开单日期','客户名称','实收金额']].groupby(by=['客户名称'],as_index=False).sum()#  数据类型问题
    Number_of_products = SalesDetails[SalesDetails['开单日期']<=deadline][['开单日期','客户名称','品名/规格']].groupby(by=['客户名称'],as_index=False).count() #品规数
    df = pd.merge(Sales,Number_of_products,on='客户名称',how='left')
    del df['开单日期']
    return df

sp = sale_product(deadline)
sp1 = sale_product(deadline1)
sp2 = sale_product(deadline2)
sp3 = sale_product(deadline3)
sp4 = sale_product(deadline4)
sale_product = pd.concat([sp,sp1],axis=1)       #截止日期



#恒兴实绩  （药品实绩）
Customer_allocation = pd.read_excel(r'F:\会员店销售数据分析表-20190714 - 副本.xlsx','7月销售品规分析汇总表（恒兴）')
Drug_name = pd.read_excel(r'F:\会员店销售数据分析表-20190714 - 副本.xlsx','目标品种目录',dtype=str)
Performance = SalesDetails[SalesDetails['商品代码'].isin(['0840900','0841000','0841100','2426400','2426401','2426402','0709901'])]

Performance = pd.merge(Performance,Customer_allocation,on='客户名称',how='left')
Performance = pd.merge(Performance,Drug_name,on='商品代码',how='left')

Performance = Performance[['最新分配','分类','客户名称']].groupby(by=['最新分配','分类'],as_index=False).count()

Store_sales = SalesDetails[SalesDetails['数量']>0]

Store_sales = pd.merge(Store_sales,Customer_allocation,on='客户名称',how='left')
Store_sales = pd.merge(Store_sales,Drug_name,on='商品代码',how='left')
Store_sales = Store_sales[['最新分配','客户名称']].drop_duplicates(['最新分配','客户名称'],keep='first').groupby(by=['最新分配'],as_index=False).count()

df4 = pd.concat([df2,Performance,Store_sales],axis=1)
df4.fillna('')

#导出excel文件
book = load_workbook(r'F:\会员店销售数据分析表-20190714 - 副本.xlsx')
writer = pd.ExcelWriter(r'F:\会员店销售数据分析表-20190714 - 副本.xlsx', engine='openpyxl')
writer.book = book
sale_product.to_excel(writer, '销售品规过渡表')
df4.to_excel(writer, '战略品种过渡表')
writer.save()