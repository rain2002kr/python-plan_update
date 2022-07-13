
from sre_constants import SUCCESS
import pandas as pd
from openpyxl import load_workbook
DEBUG_ON = 1
DEBUG_OFF = 0

debug = DEBUG_OFF
D_create_new_excel = DEBUG_OFF
D_load_time_excel = DEBUG_ON
D_load_cur_plan_excel = DEBUG_ON
D_make_df_plan_sum = DEBUG_OFF
D_make_df_plan_jobs = DEBUG_OFF

# FUNCTION CODE = read_sheets() 
# path         : 파일이 있는 경로를 지정한다. 
# excel        : path +엑셀파일.xlsx 
# return        : sheet 명을 배열 형태로 돌려준다. 
# comment      : 독립적으로 동작하고, 경로와, 엑셀파일을 넣으면 해당하는 모든 시트를 읽는다.
###############################################################################################    
def read_sheets(path, excel):
    i = -1
    wb = load_workbook(path + excel)
    shts = []
    print(excel)
    for item in wb.sheetnames:
        i = i + 1
        shts.append(item)
        print(i, item)
    return shts

# FUNCTION CODE : create_new_excel
# path = 이 안에 들어 있는 모든 엑셀 파일의 타겟 sheet를 삭제한다.
# file        : path +엑셀파일.xlsx 
# tar_sht_number : 시트번호, 첫번째 0, 두번째 1 ... n번째 n
# comment      : 독립적으로 동작하고, 엑셀파일과 시트번호을 넣으면, 해당하는 시트가 삭제된다.      
###############################################################################################    
def create_new_excel():
    path = r"D:\97. 업무공유파일\000. 계획\01. 시간분석테이블/"
    file = "03. 시간분석테이블-테스트.xlsx"
    df = pd.DataFrame()
    sht =['통합','업무','기타']
    try :
        with pd.ExcelWriter(path + file) as writer:  
            df.to_excel(writer, sheet_name= sht[0], index=False)
            df.to_excel(writer, sheet_name= sht[1], index=False)
            df.to_excel(writer, sheet_name= sht[2], index=False)
        if debug or D_create_new_excel: 
            SUCCESS = f'create_new_excel : SUCCESS'
            print(f"CODE : {SUCCESS} ") 
            read_sheets(path , file)

        
    except :
        if debug or D_create_new_excel: 
            ERROR = f'create_new_excel : ERROR 파일확장자가 지정되지않음.EX) .xlsx'
            print(f"CODE : {ERROR} ") 

# FUNCTION CODE : load_time_excel
# path = 파일이 있는 경로명
# file = 파일명.xlsx
# src  = path + file     
# 
# comment : 시간분석테이블에 내용을 읽어온다. 
###############################################################################################    
def load_time_excel():
    path = r"D:\97. 업무공유파일\000. 계획\01. 시간분석테이블/"
    file = "03. 시간분석테이블-테스트.xlsx"
    src = path + file
    try :
        df1 = pd.read_excel(src, sheet_name= 0)
        df2 = pd.read_excel(src, sheet_name= 1)
        df3 = pd.read_excel(src, sheet_name= 2)

        if debug or D_load_time_excel: 
            SUCCESS = f'load_time_excel : SUCCESS'
            print(f"CODE : {SUCCESS} ") 
        
        return df1, df2, df3
        
    except :
        if debug or D_load_time_excel: 
            ERROR = f'load_time_excel : ERROR 없는 시트에 접근 하였음'
            print(f"CODE : {ERROR} ")

# FUNCTION CODE : load_cur_plan_excel
# path = 파일이 있는 경로명
# file = 파일명.xlsx
# src  = path + file     
# cnt_sht_num : 현재날짜 부터 몇개의 시트를 읽어올것인지 파라미터
# comment : 일일계획표에 지정된 시트를 모두 읽어와서 배열로 리턴한다. 
###############################################################################################    
def load_cur_plan_excel(cmd, sht_num_cnt=1, date='220701'):
    path = r"D:\97. 업무공유파일\000. 계획/"
    file = "01. 2022 일일계획표.xlsx"
    src = path + file
    frame = []
    try :
        v_str_command = cmd
        v_int_cnt_sht_num = int(sht_num_cnt)
        v_str_date = date
        v_int_start_sht_num = 4
        v_int_tar_sht_num = v_int_start_sht_num + v_int_cnt_sht_num

        if v_str_command == 'load_shts_by_number':    
            for i in range(v_int_start_sht_num, v_int_tar_sht_num):
                df = pd.read_excel(src, sheet_name= i)
                frame.append(df)
        elif v_str_command == 'load_sht_by_date':
            df = pd.read_excel(src, sheet_name= v_str_date)
            frame.append(df)    


        if debug or D_load_cur_plan_excel: 
            SUCCESS = f'load_cur_plan_excel : SUCCESS'
            print(f"CODE : {SUCCESS} ") 
            # 디버그시 현재 날짜 출력 
            for i in range(0, len(frame)):
                print(f"\tREAD SHT DATE : {frame[i].columns[1]}") 
                
        
        return frame
        
    except :
        if debug or D_load_cur_plan_excel: 
            ERROR = f'load_cur_plan_excel : ERROR 없는 시트에 접근 하였음'
            print(f"CODE : {ERROR} ")           

# FUNCTION CODE : make_df_plan_sum
# path 
# file        : 
# comment : 
###############################################################################################    
def make_df_plan_sum(org_df):
    SUM_JOBS_FST_COL  = 11
    SUM_JOBS_END_COL  = 18
    SUM_JOBS_FST_ROW  = 21
    SUM_JOBS_END_ROW  = 22
    v_arr_sum_jobs_col_addr = []
    v_str_cur_date = org_df.columns[1]
    
    try:
        for i in range( SUM_JOBS_FST_COL, SUM_JOBS_END_COL + 1) :
            v_arr_sum_jobs_col_addr.append(i) 

        df_plan_sum = org_df.iloc[SUM_JOBS_FST_ROW : SUM_JOBS_END_ROW, v_arr_sum_jobs_col_addr].dropna()
        df_plan_sum.insert(0,"날짜", v_str_cur_date)
        
        if debug or D_make_df_plan_sum: 
            SUCCESS = f'make_df_plan_sum : SUCCESS'
            print(f"CODE : {SUCCESS} ") 
        
        return df_plan_sum
    
    except :
        if debug or D_make_df_plan_sum: 
            ERROR = f'make_df_plan_sum : ERROR DateFrame from the Sequence'
            print(f"CODE : {ERROR} ") 




def make_df_plan_jobs(org_df):
    v_arr_jobs_col_addr = []
    v_str_cur_date =  org_df.columns[1]
    JOBS_FST_COL  = 11
    JOBS_END_COL  = 24
    JOBS_STEP  = 2
    JOBS_FST_ROW  = 24
    JOBS_QTY = 7
    v_arr_jobs = []
    
    try:
        # Get jobs columns array 
        for i in range(JOBS_FST_COL, JOBS_END_COL, JOBS_STEP) :
            j = i + 1
            v_arr_jobs_col_addr.append([i,j])
        
        # Get jobs dataframe array
        for i in range(0, JOBS_QTY):
            df_jobs = org_df.iloc[ JOBS_FST_ROW:,v_arr_jobs_col_addr[i]].dropna()
            df_jobs = df_jobs.reset_index(drop=True).T.reset_index(drop=True)
            v_arr_jobs.append(df_jobs)    
        
        
        # Set jobs column name with first row value
        for i in range(0,len(v_arr_jobs)):
            v_arr_jobs[i].columns = v_arr_jobs[i].iloc[0]
            v_arr_jobs[i] = v_arr_jobs[i].drop([v_arr_jobs[i].index[0]])
        
        df_job_others = pd.concat([v_arr_jobs[1],v_arr_jobs[2],v_arr_jobs[3],v_arr_jobs[4],v_arr_jobs[5],v_arr_jobs[6]],axis=1)
        
        # insert 날짜 
        v_arr_jobs[0].insert(0, '날짜', v_str_cur_date)
        df_job_others.insert(0, '날짜', v_str_cur_date)

        if debug or D_make_df_plan_jobs: 
            SUCCESS = f'make_df_plan_jobs : SUCCESS'
            print(f"CODE : {SUCCESS} ") 

        return v_arr_jobs[0], df_job_others
    
    except :
        if debug or D_make_df_plan_jobs: 
            ERROR = f'make_df_plan_jobs : ERROR 데이터 프레임 만들기 실패'
            print(f"CODE : {ERROR} ") 
        
