
import pandas as pd
from openpyxl import load_workbook

# FUNCTION CODE : make_df_plan_sum
# path 
# file        : 
# comment : 
###############################################################################################    
def make_df_plan_sum(org_df):
    d_cols_jobs_total = []
    d_rows_jobs_total = [0]
    v_start_jobs_total = 11
    v_end_jobs_total = 18 
    
    for i in range( v_start_jobs_total, v_end_jobs_total + 1) :
        d_cols_jobs_total.append(i)
    try:
        df_plan_sum = org_df.iloc[21:22,d_cols_jobs_total].dropna()
        df_plan_sum.insert(0,"날짜", org_df.columns[1])
        return df_plan_sum
    
    except :
        ERROR = f'make_df_plan_sum : ERROR MAKE DateFrame from the Sequence'
        print(f"CODE : {ERROR} ") 


def make_df_plan_jobs(org_df):
    cur_date =  org_df.columns[1]
    d_cols_jobs_total = []
    d_rows_jobs_total = [0]
    v_start = 11
    v_end = 24 
    v_step = 2
    JOBS = []
    jobs = []
    tb_name ='tb_job'
        
    try:
        
        for i in range(v_start, v_end, v_step) :
            j = i + 1
            JOBS.append([i,j])
        
        job_name = ["업무", "계획", "개발", "필수", "펀", "교육", "이동"]
        start_row = 24

        for i in range(0,len(job_name)):
            df_jobs = org_df.iloc[start_row:,JOBS[i]].dropna()
            df_jobs = df_jobs.reset_index(drop=True)
            df_jobs_n = df_jobs.T.reset_index(drop=True)
            df_jobs_n.reset_index()
            jobs.append(df_jobs_n)    
        
        for i in range(0,len(jobs)):
            jobs[i].columns = jobs[i].iloc[0]
            jobs[i] = jobs[i].drop([jobs[i].index[0]])
            
        tdf = pd.concat([jobs[1], jobs[2],jobs[3],jobs[4],jobs[5],jobs[6]],axis=1)
        jobs[0].insert(0, '날짜', org_df.columns[1])
        tdf.insert(0, '날짜', org_df.columns[1])

        return jobs[0], tdf
    
    except :
        ERROR = f'make_df_plan_jobs : ERROR MAKE DateFrame from the Sequence'
        print(f"CODE : {ERROR} ") 
