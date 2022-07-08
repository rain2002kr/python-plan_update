from sqlalchemy import create_engine    
import pandas as pd
import openpyxl 
import numpy as np
import common_pc

sht_name = '220708'

org_df, cur_date = common_pc.make_df_from_excel(sht_name)
        
df_plan_sum = common_pc.make_df_plan_sum(org_df, cur_date)
print(df_plan_sum)

src_path = r"D:\97. 업무공유파일\000. 계획\01. 시간분석테이블/"
file =  '03. 시간분석테이블.xlsx'
df_plan_sum.to_excel(src_path+file, sheet_name = sht_name, index=False)
# common_pc.save_excel( src_path+file, sht_name = 'plan_details')


rdf = pd.read_excel(src_path+file, sheet_name = sht_name)
print('테이블 읽어오기')
print(rdf)

ndf1 = df_plan_sum.copy()
ndf1 = ndf1.iloc[0,1] = 220709

print('테이블 복사')
print(ndf1)