from dataclasses import replace
from sqlalchemy import create_engine    
import pandas as pd
from openpyxl import load_workbook
import module as md
DEBUG_ON = 1
DEBUG_OFF = 0

debug = DEBUG_OFF


D_make_df_plan = DEBUG_OFF

try:
    
# FUNCTION CODE : make_df_plan_sum
# path = 파일이 있는 경로명
# file = 파일명.xlsx
# src  = path + file     
# 
# comment :  
###############################################################################################    
    def make_df_plan(cnt_sht_num):
        path = r"D:\97. 업무공유파일\000. 계획/"
        file = "01. 2022 일일계획표.xlsx"
        src = path + file
        frame = []
        df_sum_frame = []
        df_job1_frame = []
        df_others_frame = []

        if int(cnt_sht_num) > 30 or int(cnt_sht_num) < 1 : 
            v_int_cnt_sht_num = 1
        else : 
            v_int_cnt_sht_num = cnt_sht_num


        

        try :
            # get 계획표 from 일일계획표, 현재날짜 + 갯수 
            frame = md.load_cur_plan_excel(v_int_cnt_sht_num)
            
            for i in range(0, len(frame)):
                df_plan_sum = md.make_df_plan_sum(frame[i])
                df_plan_job1, df_plan_others = md.make_df_plan_jobs(frame[i])
                df_sum_frame.append(df_plan_sum)
                df_job1_frame.append(df_plan_job1)
                df_others_frame.append(df_plan_others)
            

            if debug or D_make_df_plan: 
                SUCCESS = f'make_df_plan : SUCCESS'
                print(f"CODE : {SUCCESS} ")
                print(f"\tREAD SUM FRAME : {df_sum_frame}")     
                print(f"\tREAD SUM FRAME : {df_job1_frame}")     
                print(f"\tREAD SUM FRAME : {df_others_frame}")     
            
        except :
            if debug or D_make_df_plan: 
                ERROR = f'make_df_plan : ERROR 없는 시트에 접근 하였음'
                print(f"CODE : {ERROR} ")        


    
    #######################################################################################################
    # MAIN CODE START 
    
    src_path = r"D:\000. 발주서\02. SRC_M/"
    del_sht_name = 'Order'
    del_sht_number = 1
    del_sht_start_number = 1
    del_sht_end_number = 2
    SHT_NAME_TYPE = 0
    SHT_NUM_TYPE = 1
    SHT_RNG_TYPE = 2
    SHT_READ_TYPE = 3
    READ_DFE = 1
    v_int_read_sht_qty = 11

    # md.create_new_excel()
    # md.load_time_excel()            
    # md.load_cur_plan_excel()
    # md.make_df_plan()
    # make_df_plan(v_int_read_sht_qty)

except Exception as ex:
    print('error' + str(ex))
    