from dataclasses import replace
from sqlalchemy import create_engine
import pandas as pd
from openpyxl import load_workbook
import module as md
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import ui_module as u_md


def main_process(self, cmd, date="220701", sht_num_cnt=0):
    v_str_command = cmd
    v_int_date = date
    v_int_sht_num_cnt = int(sht_num_cnt)

    if v_str_command == "init":
        print("init ")
        bar_timer_on(self, "start", 10)
        md.create_new_excel()

    elif v_str_command == "save_server":
        print("save_server ")

        bar_timer_on(self, "start", 50)
        md.save_time_excel_to_server()

    elif v_str_command == "load_server":
        print("load_server ")
        df_jobs = [[], [], []]
        df_jobs[0], df_jobs[1], df_jobs[2] = md.load_time_excel_from_server()

        md.save_time_excel(df_jobs[0], df_jobs[1], df_jobs[2])
        bar_timer_on(self, "start", 10)
        self.tbw = u_md.set_df_table_2_arr(self, df_jobs)

    elif v_str_command == "time_load":
        print("time_load")
        bar_timer_on(self, "start", 10)
        jobs = [[], [], []]
        jobs[0], jobs[1], jobs[2] = md.load_time_excel()
        self.tbw = u_md.set_df_table_2_arr(self, jobs)

    elif v_str_command == "load":
        print("load")
        bar_timer_on(self, "start", 10)
        jobs = [[], [], []]
        if v_int_sht_num_cnt > 1:
            jobs[0], jobs[1], jobs[2] = md.make_df_plan(
                "load_shts_by_number", v_int_sht_num_cnt, v_int_date
            )
        else:
            jobs[0], jobs[1], jobs[2] = md.make_df_plan(
                "load_sht_by_date", v_int_sht_num_cnt, v_int_date
            )
        self.tbw = u_md.set_df_table_3_arr(self, jobs)

    elif v_str_command == "save":
        print("save")
        bar_timer_on(self, "start", 10)
        jobs = [[], [], []]
        if v_int_sht_num_cnt > 1:
            jobs[0], jobs[1], jobs[2] = md.make_df_plan(
                "load_shts_by_number", v_int_sht_num_cnt, v_int_date
            )
        else:
            jobs[0], jobs[1], jobs[2] = md.make_df_plan(
                "load_sht_by_date", v_int_sht_num_cnt, v_int_date
            )

        df_jobs = md.update_df_from_arr(jobs)

        md.save_time_excel(df_jobs[0], df_jobs[1], df_jobs[2])

        self.tbw = u_md.set_df_table_3_arr(self, jobs)

    elif v_str_command == "scr_clear":
        print("scr_clear")
        bar_timer_on(self, "start", 5)
        self.tbw[0].clear()
        self.tbw[1].clear()
        self.tbw[2].clear()

    else:
        print("jobs 값이 없음")


def bar_timer_on(self, commend, base_time=10):
    if commend == "init":
        self.bar1_timer.timeout.connect(lambda: timerEvent(self, 1))
        self.bar1.setValue(0)
        self.bar1_time = QTime(0, 0, 0)

    elif commend == "start":
        self.bar1_timer.timeout.connect(lambda: timerEvent(self, 1))
        self.bar1.setValue(0)
        self.bar1_time = QTime(0, 0, 0)
        self.bar1_timer.start(base_time)

    elif commend == "end":
        self.bar1_timer.stop()
    else:
        print("timer 동작안함")


def timerEvent(self, time_interval=1):
    self.bar1_time = self.bar1_time.addSecs(time_interval)
    v_int_time = int(self.bar1_time.toString("ss"))
    self.bar1.setValue(v_int_time)

    if v_int_time >= self.bar1.maximum():
        self.bar1_timer.stop()
        return


# DEBUG_ON = 1
# DEBUG_OFF = 0

# debug = DEBUG_OFF
# D_make_df_plan = DEBUG_OFF

# try:

# # FUNCTION CODE : make_df_plan_sum
# # path = 파일이 있는 경로명
# # file = 파일명.xlsx
# # src  = path + file
# #
# # comment :
# ###############################################################################################
#     def make_df_plan(cmd, cnt_sht_num=1, date='220701'):
#         v_str_command = cmd
#         v_int_cnt_sht_num = cnt_sht_num
#         v_str_sht_date = date

#         path = r"D:\97. 업무공유파일\000. 계획/"
#         file = "01. 2022 일일계획표.xlsx"
#         src = path + file
#         frame = []
#         df_sum_frame = []
#         df_job1_frame = []
#         df_others_frame = []

#         try :
#             # get 계획표 from 일일계획표, 현재날짜 + 갯수
#             if v_str_command == 'load_shts_by_number':
#                 frame = md.load_cur_plan_excel('load_shts_by_number', v_int_cnt_sht_num, v_str_sht_date)
#             else:
#                 frame = md.load_cur_plan_excel('load_sht_by_date', v_int_cnt_sht_num, v_str_sht_date)

#             for i in range(0, len(frame)):
#                 df_plan_sum = md.make_df_plan_sum(frame[i])
#                 df_plan_job1, df_plan_others = md.make_df_plan_jobs(frame[i])
#                 df_sum_frame.append(df_plan_sum)
#                 df_job1_frame.append(df_plan_job1)
#                 df_others_frame.append(df_plan_others)

#             if debug or D_make_df_plan:
#                 SUCCESS = f'make_df_plan : SUCCESS'
#                 print(f"CODE : {SUCCESS} ")
#                 print(f"LOAD BY: {v_str_command} ")
#                 print(f"\tREAD SUM FRAME : {df_sum_frame}")
#                 print(f"\tREAD SUM FRAME : {df_job1_frame}")
#                 print(f"\tREAD SUM FRAME : {df_others_frame}")

#             return df_sum_frame, df_job1_frame, df_others_frame

#         except :
#             if debug or D_make_df_plan:
#                 ERROR = f'make_df_plan : ERROR 없는 시트에 접근 하였음'
#                 print(f"CODE : {ERROR} ")


#     #######################################################################################################
#     # MAIN CODE START

#     src_path = r"D:\000. 발주서\02. SRC_M/"
#     del_sht_name = 'Order'
#     del_sht_number = 1
#     del_sht_start_number = 1
#     del_sht_end_number = 2
#     SHT_NAME_TYPE = 0
#     SHT_NUM_TYPE = 1
#     SHT_RNG_TYPE = 2
#     SHT_READ_TYPE = 3
#     READ_DFE = 1
#     v_int_read_sht_qty = 11

#     # md.create_new_excel()
#     # md.load_time_excel()
#     # md.load_cur_plan_excel()
#     # md.make_df_plan()
#     # make_df_plan(v_int_read_sht_qty)

# except Exception as ex:
#     print('error' + str(ex))
