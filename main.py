import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QHBoxLayout, QVBoxLayout, QTableWidgetItem, QTableWidget, QGridLayout, QProgressBar, QLabel, QSpinBox, QHeaderView, QSizePolicy
from PyQt5.QtCore import QCoreApplication, QBasicTimer, Qt, QTimer, QTime,QSize
import pandas as pd

import process as ps
import module as md
import ui_module as u_md


DEBUG_ON = 1
DEBUG_OFF = 0

debug = DEBUG_OFF
D_main_process = DEBUG_OFF

### 사이즈 정책을 설정한 새로운 class를 생성합니다. ###
class QPushButton(QPushButton):
    def __init__(self, parent = None):
        super().__init__(parent)
        # self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # self.setStyleSheet("QPushButton{color: white;\n"
        #             "background-color:qlineargradient(spread:reflect, x1:1, y1:0, x2:0.995, y2:1, stop:0 rgba(218, 218, 218, 255), stop:0.305419 rgba(0, 7, 11, 255), stop:0.935961 rgba(2, 11, 18, 255), stop:1 rgba(240, 240, 240, 255));\n"
        #             "border: 1px solid black;\n"
        #             "border-radius: 20px;}\n"
        # )
        
        

# class QLabel(QLabel):
#     def __init__(self, parent = None):
#         super().__init__(parent)
#         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

# class QLineEdit(QLineEdit):
#     def __init__(self, parent = None):
#         super().__init__(parent)
#         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class Exam(QWidget):
    def __init__(self):
        super().__init__()
        self.defineItem()
        self.Decolation()
        self.initUI()
        self.UIconnectFC()
    
    
    def defineItem(self):
        self.lb_subject = QLabel('Hoons PLAN Daily update')
        # self.btn_init = QPushButton('init', self)
        self.btn_init = QPushButton('init')
        self.btn_save_server = QPushButton('save_server')
        self.btn_load_server = QPushButton('load_server')
        self.bar1 = QProgressBar(self)
        self.bar1.setOrientation(Qt.Horizontal)
        self.bar1.setRange(0,30)

        self.bar1_timer = QTimer()
        self.bar1_time = QTime(0,0,0)

        self.btn_timeload = QPushButton('timeload')
        self.btn_load = QPushButton('load')
        self.btn_save = QPushButton('save')
        self.btn_scr_clr = QPushButton('scr_clear')
        self.ed_date = QLineEdit('220701')
        self.spin_date = QSpinBox()
        
        self.hbox1 = QHBoxLayout()
        self.hbox2 = QHBoxLayout()
        self.hbox3 = QHBoxLayout()

        self.hbox1.addWidget(self.lb_subject)
        self.hbox2.addWidget(self.btn_init)
        self.hbox2.addWidget(self.btn_save_server)
        self.hbox2.addWidget(self.btn_load_server)

        self.hbox2.addWidget(self.bar1)
        self.hbox3.addWidget(self.btn_timeload)
        self.hbox3.addWidget(self.btn_load)
        self.hbox3.addWidget(self.btn_save)
        self.hbox3.addWidget(self.btn_scr_clr)
        self.hbox3.addWidget(self.ed_date)
        self.hbox3.addWidget(self.spin_date)
        
    def Decolation(self):
        self.bar1.setStyleSheet("QProgressBar{\n"
                "    background-color: rgb(98, 114, 164);\n"
                "    color:rgb(200,200,200);\n"
                "    border-style: none;\n"
                "    border-bottom-right-radius: 10px;\n"
                "    border-bottom-left-radius: 10px;\n"
                "    border-top-right-radius: 10px;\n"
                "    border-top-left-radius: 10px;\n"
                "    text-align: center;\n"
                "}\n""QProgressBar::chunk{\n"
                "    border-bottom-right-radius: 10px;\n"
                "    border-bottom-left-radius: 10px;\n"
                "    border-top-right-radius: 10px;\n"
                "    border-top-left-radius: 10px;\n"
                "    background-color: qlineargradient(spread:pad, x1:0, y1:0.511364, x2:1, y2:0.523, stop:0 rgba(254, 121, 199, 255), stop:1 rgba(170, 85, 255, 255));\n"
                "}\n"
                "\n"
                "")
        self.btn_init.setMaximumHeight(200)
        # self.btn_init.setStyleSheet("QPushButton{color: white;\n"
        #             "background-color:qlineargradient(spread:reflect, x1:1, y1:0, x2:0.995, y2:1, stop:0 rgba(218, 218, 218, 255), stop:0.305419 rgba(0, 7, 11, 255), stop:0.935961 rgba(2, 11, 18, 255), stop:1 rgba(240, 240, 240, 255));\n"
        #             "border: 1px solid black;\n"
        #             "border-radius: 20px;}\n"
        # )
        # self.btn_save_server.setStyleSheet("QPushButton{color: white;\n"
        #             "background-color: rgb(58, 134, 255);\n"
        #             "border-radius: 5px;\n}")
        
    def UIconnectFC(self):
        self.btn_init.clicked.connect(lambda x: self.main_process('init'))
        self.btn_save_server.clicked.connect(lambda x: self.main_process('save_server'))
        self.btn_load_server.clicked.connect(lambda x: self.main_process('load_server'))

        self.btn_timeload.clicked.connect(lambda x: self.main_process('time_load'))
        self.btn_load.clicked.connect(lambda x: self.main_process('load', self.ed_date.text(), self.spin_date.text()))
        self.btn_save.clicked.connect(lambda x: self.main_process('save', self.ed_date.text(), self.spin_date.text()))
        self.btn_scr_clr.clicked.connect(lambda x: self.main_process('scr_clear'))



    def initUI(self):
        self.tbw = {}
        for i in range(0,3):
            self.tbw[i] = QTableWidget()
            
        # self.btn_init.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # self.btn_save_server.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # self.btn_load_server.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # layout = QVBoxLayout()
        self.spin_date.setMaximum(30)
        self.spin_date.setMinimum(0)
        self.spin_date.setSingleStep(1)
        

        layout = QGridLayout()
        layout.setSpacing(10)

        # layout.addLayout(self.hbox1 ,0,0)
        layout.addLayout(self.hbox2 ,1,0)
        layout.addLayout(self.hbox3 ,2,0)
        layout.addWidget(self.tbw[0],3,0)
        layout.addWidget(self.tbw[1],4,0)
        layout.addWidget(self.tbw[2],5,0)
        
        self.setLayout(layout)
        
        self.setWindowTitle('하루 기록 저장프로그램')
        self.setGeometry(0,0,500,500)
        self.show()


    def main_process(self, cmd, date= '220701', sht_num_cnt = 0):
        v_str_command = cmd
        v_int_date = date
        v_int_sht_num_cnt = int(sht_num_cnt)

        if v_str_command == 'init':
            print("init ")
            self.bar_timer_on("start",10)
            md.create_new_excel()
        
        elif v_str_command == 'save_server':
            print("save_server ")
            
            self.bar_timer_on("start",50)
            md.save_time_excel_to_server()

            

        
        elif v_str_command == 'load_server':
            print("load_server ")
            df_jobs=[[],[],[]]
            df_jobs[0], df_jobs[1], df_jobs[2] = md.load_time_excel_from_server()

            md.save_time_excel(df_jobs[0], df_jobs[1], df_jobs[2])
            self.bar_timer_on("start",10)
            self.tbw = u_md.set_df_table_2_arr(self, df_jobs)
            

        elif v_str_command == 'time_load':
            print('time_load')
            self.bar_timer_on("start",10)
            jobs=[[],[],[]]
            jobs[0], jobs[1], jobs[2] = md.load_time_excel()
            self.tbw = u_md.set_df_table_2_arr(self, jobs)


        elif v_str_command == 'load':
            print('load')
            self.bar_timer_on("start",10)
            jobs =[[],[],[]]
            if v_int_sht_num_cnt > 1: 
                jobs[0] ,jobs[1], jobs[2] = ps.make_df_plan('load_shts_by_number', v_int_sht_num_cnt,  v_int_date)
            else :
                jobs[0], jobs[1], jobs[2] = ps.make_df_plan('load_sht_by_date', v_int_sht_num_cnt,  v_int_date)
            self.tbw = u_md.set_df_table_3_arr(self, jobs)
            

        elif v_str_command == 'save':
            print('save')
            self.bar_timer_on("start",10)
            jobs =[[],[],[]]
            if v_int_sht_num_cnt > 1: 
                jobs[0] ,jobs[1], jobs[2] = ps.make_df_plan('load_shts_by_number', v_int_sht_num_cnt,  v_int_date)
            else :
                jobs[0], jobs[1], jobs[2] = ps.make_df_plan('load_sht_by_date', v_int_sht_num_cnt,  v_int_date)
            
            df_jobs = md.update_df_from_arr(jobs)

            md.save_time_excel(df_jobs[0], df_jobs[1], df_jobs[2])

            self.tbw = u_md.set_df_table_3_arr(self, jobs)

        elif v_str_command == 'scr_clear':
            print('scr_clear')
            self.bar_timer_on("start",5)
            self.tbw[0].clear()
            self.tbw[1].clear()
            self.tbw[2].clear()
            
        

        else:
            print("jobs 값이 없음")


    def bar_timer_on(self, commend, base_time = 10):
        if commend == "init": 
            self.bar1_timer.timeout.connect(lambda : self.timerEvent(1))  
            self.bar1.setValue(0)
            self.bar1_time = QTime(0,0,0) 

        elif commend == "start":
            self.bar1_timer.timeout.connect(lambda : self.timerEvent(1))  
            self.bar1.setValue(0)
            self.bar1_time = QTime(0,0,0) 
            self.bar1_timer.start(base_time)


        elif commend == "end": 
            self.bar1_timer.stop()
        else :
            print("timer 동작안함")


    def timerEvent(self, time_interval = 1):
        self.bar1_time = self.bar1_time.addSecs(time_interval)
        v_int_time = int(self.bar1_time.toString("ss"))
        self.bar1.setValue(v_int_time)
        
        if v_int_time >= self.bar1.maximum():
            self.bar1_timer.stop()
            return 
            

app = QApplication(sys.argv)
w = Exam()

# 윈도우창에 이벤트처리정보를 위젯 객체에 넘겨준다. 
# 메인 루프라고 한다. app.exec_() 끝나야 sys.exit 함수가 호출된다. 
sys.exit(app.exec_())

