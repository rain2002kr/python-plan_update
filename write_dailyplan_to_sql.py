from dataclasses import replace
from sqlalchemy import create_engine    
import pandas as pd
from openpyxl import load_workbook
import common_pc


# 엑셀 저장 리팩터링 필요 
def main(start_date, cnt):
    engine = common_pc.connect_sql_server()

    if cnt == 0:
        date_b = '22'+str(start_date)
        sht_name = date_b 
        org_df, cur_date = common_pc.make_df_from_excel(sht_name)
        print(org_df)
        
        
        df_plan_sum = common_pc.make_df_plan_sum(engine, org_df, cur_date)
        print(df_plan_sum)

        jobs = common_pc.make_df_plan_jobs(engine, org_df, cur_date)
        print(jobs)

        tdf = pd.concat([jobs[1], jobs[2],jobs[3],jobs[4],jobs[5],jobs[6]],axis=1)
        jobs[0].insert(0, '날짜', org_df.columns[1])
        tdf.insert(0, '날짜', org_df.columns[1])

        print(tdf)
        print(jobs[0])
        common_pc.save_excel( tdf, sht_name = 'plan_details')
        common_pc.save_excel( jobs[0], sht_name = 'plan_works')

        ndf = common_pc.read_df_from_timetable('plan_works')
        print(ndf)
        ndf = common_pc.read_df_from_timetable('plan_details')
        print(ndf)


# def load_timesheet():
#     pd.read_excel()

def load_cur_excel(date):
    sht_name = str(date) 
    org_df, cur_date = common_pc.make_df_from_excel(sht_name)
    # print(org_df)
    
    df_plan_sum = common_pc.make_df_plan_sum(org_df, cur_date)
    # print(df_plan_sum)

    job0, tdf = common_pc.make_df_plan_jobs(org_df, cur_date)
    # print(job0)
    # print(tdf)
    return df_plan_sum, job0, tdf
    

# start_date = '0702'
# cnt = 0
# main(start_date, cnt)


