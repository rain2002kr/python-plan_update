from dataclasses import replace
from sqlalchemy import create_engine
import pandas as pd
from openpyxl import load_workbook
import module as md
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import ui_module as u_md


def main_process(self, cmd, date="230101", sht_num_cnt=0):
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
