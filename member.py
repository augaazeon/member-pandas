import pandas as pd 
import numpy as np
from openpyxl import load_workbook
SalesDetails = pd.read_excel(r'C:\Users\20245\Desktop\��Ա��0718\����7.1-7.14������ϸ.xlsx',dtype=str)

#��������
SalesVolume = SalesDetails[['��Ʒ����','����']]
SalesVolume['����'] = SalesVolume['����'].astype('int')
SalesVolume = SalesVolume.groupby(by=['��Ʒ����'],as_index=False).sum()#  ������������
#�����ͻ���
Number_of_customers = SalesDetails[['�ͻ�����','��Ʒ����']]
Number_of_customers = Number_of_customers.drop_duplicates(['�ͻ�����','��Ʒ����'],keep='first')  #ok
Number_of_customers = Number_of_customers.groupby(by=['��Ʒ����'],as_index=False).count()
#����Ƶ��
Frequency_of_sales = SalesDetails[['��������','��Ʒ����']]
Frequency_of_sales = Frequency_of_sales.drop_duplicates(['��������','��Ʒ����'],keep='first') 
Frequency_of_sales = Frequency_of_sales.groupby(by=['��Ʒ����'],as_index=False).count()
df1 = pd.merge(SalesVolume,Number_of_customers,on='��Ʒ����',how='left')
df2 = pd.merge(df1,Frequency_of_sales,on='��Ʒ����',how='left')  #���ɱ��ṩ����ս��Ʒ�֡�


#���۶�
deadline = pd.to_datetime('2019-07-07')#��ֹ����
deadline1 = pd.to_datetime('2019-07-14')
deadline2 = pd.to_datetime('2019-07-21')
deadline3 = pd.to_datetime('2019-07-28')
deadline4 = pd.to_datetime('2019-07-31')

SalesDetails['��������'] = pd.to_datetime(SalesDetails['��������'])
Sales = SalesDetails[SalesDetails['��������']<=deadline]
Sales = Sales[['��������','�ͻ�����','ʵ�ս��']]
Sales.groupby(by=['�ͻ�����'],as_index=False).sum()#  ������������
Sales['ʵ�ս��'] = Sales['ʵ�ս��'].str.replace(',', '')#���Ҹ�ʽת��
Sales['ʵ�ս��'] = Sales['ʵ�ս��'].astype('float64')#������
Sales = Sales.groupby(by=['�ͻ�����'],as_index=False).sum()
#Ʒ����
Number_of_products = SalesDetails[SalesDetails['��������']<=deadline]
Number_of_products = Number_of_products[['��������','�ͻ�����','Ʒ��/���']]
Number_of_products = Number_of_products.groupby(by=['�ͻ�����'],as_index=False).count()

df3 = pd.merge(Sales,Number_of_products,on='�ͻ�����',how='left')  #���ɱ��ṩ��Ʒ�����
del df3['��������']


#����ʵ��  ��ҩƷʵ����
Customer_allocation = pd.read_excel(r'C:\Users\20245\Desktop\��Ա���������ݷ�����-20190714 - ����.xlsx','7������Ʒ��������ܱ����ˣ�')
Drug_name = pd.read_excel(r'C:\Users\20245\Desktop\��Ա���������ݷ�����-20190714 - ����.xlsx','Ŀ��Ʒ��Ŀ¼',dtype=str)
Performance = SalesDetails[SalesDetails['��Ʒ����'].isin(['0840900','0841000','0841100','2426400','2426401','2426402','0709901'])]
Performance = pd.merge(Performance,Customer_allocation,on='�ͻ�����',how='left')
Performance = pd.merge(Performance,Drug_name,on='��Ʒ����',how='left')
Performance = Performance[['���·���','����','�ͻ�����']]
#Performance['����'] = Performance['����'].astype('int') 
Performance = Performance.groupby(by=['���·���','����'],as_index=False).count()

SalesDetails['����'] = SalesDetails['����'].astype(int)
Store_sales = SalesDetails[SalesDetails['����']>0]
Store_sales = pd.merge(Store_sales,Customer_allocation,on='�ͻ�����',how='left')
Store_sales = pd.merge(Store_sales,Drug_name,on='��Ʒ����',how='left')
Store_sales = Store_sales[['���·���','�ͻ�����']]
Store_sales = Store_sales.drop_duplicates(['���·���','�ͻ�����'],keep='first')
Store_sales = Store_sales.groupby(by=['���·���'],as_index=False).count()

df4 = pd.concat([df2,df3,Performance,Store_sales],axis=1)
df4.fillna('')

#����excel�ļ�
book = load_workbook(r'C:\Users\20245\Desktop\��Ա���������ݷ�����-20190714 - ����.xlsx')
writer = pd.ExcelWriter(r'C:\Users\20245\Desktop\��Ա���������ݷ�����-20190714 - ����.xlsx', engine='openpyxl')
writer.book = book

df4.to_excel(writer, '���ɱ�')
writer.save()