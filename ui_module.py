from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import pandas as pd
from openpyxl import load_workbook
import module as md
DEBUG_ON = 1
DEBUG_OFF = 0

debug = DEBUG_OFF
D_set_df_table_3_arr = DEBUG_OFF
D_set_df_table_2_arr = DEBUG_OFF

try:
    
# FUNCTION CODE : 
# 
# 
# 
# 3 array 타입 
# comment :  PLAN 에 통합시트시간, 업무시간, 학습 및 나머지 시간을 읽어와서 TABLE 에 뿌려준다. 
###############################################################################################    
    def set_df_table_3_arr(self, jobs):
        try :
            for k in range(0,len(self.tbw)):
                for i in range(0,len(jobs[k])):
                    self.tbw[k].setRowCount(len(jobs[k]))
                    self.tbw[k].setColumnCount(len(jobs[k][i].columns))
                    self.tbw[k].setHorizontalHeaderLabels(jobs[k][i].columns)
        
            for i in range(0,len(jobs)):
                    # print(f'START {l}')
                    for j in range(0,len(jobs[i])):
                        # print(f'I V {i}')
                        for k in range(0,len(jobs[i][j].columns)):
                            # print(f'J {j}')
                            # print(jobs[i][j].iloc[0,k])
                            self.tbw[i].setItem(j,k,QTableWidgetItem(str(jobs[i][j].iloc[0, k])))    
            if debug or D_set_df_table_3_arr: 
                SUCCESS = f'set_df_table_3_arr : SUCCESS'
                print(f"CODE : {SUCCESS} ")
                print(f"\tREAD JOBS FRAME : {jobs}")     
            
            return self.tbw        
        except :
            if debug or D_set_df_table_3_arr: 
                ERROR = f'set_df_table_3_arr : ERROR'
                print(f"CODE : {ERROR} ")        
    
    
# FUNCTION CODE : 
# 
# 
# 
# 2 array 타입 
# comment :  PLAN 에 통합시트시간, 업무시간, 학습 및 나머지 시간을 읽어와서 TABLE 에 뿌려준다. 
###############################################################################################    
    def set_df_table_2_arr(self, jobs):
        try :
            for k in range(0,len(self.tbw)):
                self.tbw[k].setRowCount(len(jobs[k].index))
                self.tbw[k].setColumnCount(len(jobs[k].columns))
                self.tbw[k].setHorizontalHeaderLabels(jobs[k].columns)
        
                for i in range(len(jobs[k].index)):
                        for j in range(len(jobs[k].columns)):
                            self.tbw[k].setItem(i,j,QTableWidgetItem(str(jobs[k].iloc[i, j])))

            if debug or D_set_df_table_2_arr: 
                SUCCESS = f'set_df_table_2_arr : SUCCESS'
                print(f"CODE : {SUCCESS} ")
                print(f"\tREAD JOBS FRAME : {jobs}")     
            
            return self.tbw        
        except :
            if debug or D_set_df_table_2_arr: 
                ERROR = f'set_df_table_2_arr : ERROR'
                print(f"CODE : {ERROR} ")
    
                    
except Exception as ex:
    print('error' + str(ex))
    



