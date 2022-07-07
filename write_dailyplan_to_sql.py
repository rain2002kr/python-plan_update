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

        # tb_name = 'tb_plan_sum'
        # df1_sql_sum = common_pc.make_df_from_sql(engine, tb_name, cur_date)
        # print(df1_sql_sum)
        # src_path = r"D:\97. 업무공유파일\000. 계획/"
        # tar_path = src_path
        # file =  '03. 시간분석테이블.xlsx'
        # df = df1_sql_sum
        # sht_name = tb_name
        # common_pc.save_excel(src_path, tar_path , file , df, sht_name)
        

        


        # tb_name = 'tb_job'
        # df_jobs = common_pc.read_dfs_from_sql(engine, tb_name , cur_date)
        # print(df_jobs)
        

    # else:
    #     # 반복문 업데이트 필요 
    #     date_b = '22'+str(start_date)
    #     date_buf = []
    #     for i in range(5,6):
    #         date_buf.append(date_b + str(i))
    #         sht_name = date_b + str(i)
    #         org_df, cur_date = common_pc.make_df_from_excel(sht_name)
    #         print(org_df)
            
    #         df_plan_sum = common_pc.make_df_plan_sum(engine, org_df, cur_date)
    #         print(df_plan_sum)

    #         tb_name = 'tb_plan_sum'
    #         df1_sql_sum = common_pc.make_df_from_sql(engine, tb_name, cur_date)
    #         print(df1_sql_sum)
            

    #         jobs = common_pc.make_df_plan_jobs(engine, org_df, cur_date)
    #         print(jobs)

    #         tb_name = 'tb_job'
    #         df_jobs = common_pc.read_dfs_from_sql(engine, tb_name , cur_date)
    #         print(df_jobs)

start_date = '0702'
cnt = 0
main(start_date, cnt)


