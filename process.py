from dataclasses import replace
from sqlalchemy import create_engine    
import pandas as pd
from openpyxl import load_workbook
import module as md
DEBUG_ON = 1
DEBUG_OFF = 0

debug = DEBUG_OFF
debug_create_new_excel = DEBUG_OFF
debug_load_time_excel = DEBUG_OFF
debug_load_cur_plan_excel = DEBUG_OFF
make_df_plan = DEBUG_ON

try:
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
            if debug or debug_create_new_excel: 
                print('CODE LINE : create_new_excel OK')
            
        except :
            if debug or debug_create_new_excel: 
                print('ERROR CODE : create_new_excel()')
                print('ERROR : 파일확장자가 지정되지않음.EX) .xlsx')

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

            if debug or debug_load_time_excel: 
                print('CODE LINE : load_time_excel OK')
            
            return df1, df2, df3
            
        except :
            if debug or debug_load_time_excel: 
                print('ERROR CODE : load_time_excel()')
                print('ERROR : 없는 시트에 접근 하였음.')
            
    # FUNCTION CODE : load_cur_plan_excel
    # path = 파일이 있는 경로명
    # file = 파일명.xlsx
    # src  = path + file     
    # 
    # comment : 일일계획표에 지정된 시트를 모두 읽어와서 배열로 리턴한다. 
    ###############################################################################################    
    def load_cur_plan_excel():
        path = r"D:\97. 업무공유파일\000. 계획/"
        file = "01. 2022 일일계획표.xlsx"
        src = path + file
        frame = []
        try :
            
            cnt_sht_num = 9
            start_sht_num = 4
            tar_sht_num = start_sht_num + cnt_sht_num
            for i in range(start_sht_num, tar_sht_num):
                df = pd.read_excel(src, sheet_name= i)
                frame.append(df)


            if debug or debug_load_cur_plan_excel: 
                for i in range(0, len(frame)):
                    print(frame[i].columns[1])    
                print('CODE LINE : load_cur_plan_excel OK')
            
            
            return frame
            
        except :
            if debug or debug_load_cur_plan_excel: 
                print('ERROR CODE : load_cur_plan_excel()')
                print('ERROR : 없는 시트에 접근 하였음.')       
    
    # FUNCTION CODE : make_df_plan_sum
    # path = 파일이 있는 경로명
    # file = 파일명.xlsx
    # src  = path + file     
    # 
    # comment :  
    ###############################################################################################    
    def make_df_plan():
        path = r"D:\97. 업무공유파일\000. 계획/"
        file = "01. 2022 일일계획표.xlsx"
        src = path + file
        frame = []
        try :
            frame = load_cur_plan_excel()
            # for i in range(0, len(frame)):
            #     print(frame[i].columns[1])    
            
            # print(frame[0])
            df_plan_sum = md.make_df_plan_sum(frame[0])
            print(df_plan_sum)    
            df_plan_job, tdf = md.make_df_plan_jobs(frame[0])
            print(df_plan_sum)    


            if debug or make_df_plan: 
                print('CODE LINE : make_df_plan OK')
            
            
        except :
            if debug or make_df_plan_sum: 
                print('ERROR CODE : make_df_plan()')
                print('ERROR : 없는 시트에 접근 하였음.')       

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

    # create_new_excel()
    # load_time_excel()
    # load_cur_plan_excel()
    make_df_plan()

except Exception as ex:
    print('error' + str(ex))
