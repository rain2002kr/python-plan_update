import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QHBoxLayout, QVBoxLayout, QTableWidgetItem, QTableWidget, QGridLayout, QProgressBar, QLabel
from PyQt5.QtCore import QCoreApplication, QBasicTimer, Qt
from write_dailyplan_to_sql import main ,load_cur_excel ,load_time_excel, save_time_excel
import common_pc
import pandas as pd



class Exam(QWidget):
    def __init__(self):
        super().__init__()
        self.defineItem()
        self.initUI()
        self.functionCode()
    
    
    def defineItem(self):
        self.lb_subject = QLabel('Hoons PLAN Daily update')
        self.btn_init = QPushButton('init', self)
        self.btn_save_server = QPushButton('save_server', self)
        self.btn_load_server = QPushButton('load_server', self)
        self.bar1 = QProgressBar(self)
        self.bar1.setOrientation(Qt.Horizontal)

        self.btn_timeload = QPushButton('timeload', self)
        self.btn_load = QPushButton('load', self)
        self.btn_save = QPushButton('save', self)
        self.ed_date = QLineEdit('220702', self)
        self.ed2_date = QLineEdit('', self)
        
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
        self.hbox3.addWidget(self.ed_date)
        self.hbox3.addWidget(self.ed2_date)
        

        
    def functionCode(self):
        self.btn_init.clicked.connect(lambda x: self.main_process('init'))
        self.btn_timeload.clicked.connect(lambda x: self.main_process('time_load'))
        self.btn_load.clicked.connect(lambda x: self.main_process('load'+ self.ed_date.text()))
        self.btn_save.clicked.connect(lambda x: self.main_process('save'+ self.ed_date.text(), self.ed2_date.text()))

    def initUI(self):
        self.tbw = {}
        for i in range(0,3):
            self.tbw[i] = QTableWidget()    

        # layout = QVBoxLayout()
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

    def main_process(self, cmd, sht_num_cnt = 1):
        command = cmd[0:4]
        date = cmd[4:]
        print(command)
        print(date)
        
        if cmd == 'time_load':
            jobs=[[],[],[]]
            jobs[0], jobs[1], jobs[2] = load_time_excel()
            
            for i in range(0,2):
                if not jobs[i].empty:
                    print("jobs 값이 있음")
                    print(jobs[i])
                else:
                    print("jobs 값이 없음")
                    print(jobs[i])

            for k in range(0,len(self.tbw)):
                self.tbw[k].setRowCount(len(jobs[k].index))
                self.tbw[k].setColumnCount(len(jobs[k].columns))
                self.tbw[k].setHorizontalHeaderLabels(jobs[k].columns)
                for i in range(len(jobs[k].index)):
                    for j in range(len(jobs[k].columns)):
                        self.tbw[k].setItem(i,j,QTableWidgetItem(str(jobs[k].iloc[i, j])))
            
        if command == 'load':
            jobs=[[],[],[]]
            jobs[0], jobs[1], jobs[2] = load_cur_excel(date)
            
            for k in range(0,len(self.tbw)):
                self.tbw[k].setRowCount(len(jobs[k].index))
                self.tbw[k].setColumnCount(len(jobs[k].columns))
                self.tbw[k].setHorizontalHeaderLabels(jobs[k].columns)
                for i in range(len(jobs[k].index)):
                    for j in range(len(jobs[k].columns)):
                        self.tbw[k].setItem(i,j,QTableWidgetItem(str(jobs[k].iloc[i, j])))

        if command == 'save':
            newjobs=[[],[],[]]
            jobs=[[],[],[]]
            sumjobs = []
            newjobs[0], newjobs[1], newjobs[2] = load_cur_excel(date)
            jobs[0], jobs[1], jobs[2] = load_time_excel()
            

            if not jobs[0].empty:
                print("jobs 값이 있음")
                for i in range(0,3):
                    sumjob = pd.concat([jobs[i],newjobs[i]])
                    print("CODE : SUMJOB")
                    print(sumjob)
                    sumjobs.append(sumjob)
                save_time_excel(sumjobs[0],sumjobs[1],sumjobs[2])
        if cmd == 'init':
            print("init ")
        
        if cmd == 'save_server':
            print("save_server ")
        
        if cmd == 'load_server':
            print("load_server ")

        else:
            print("jobs 값이 없음")
            save_time_excel(newjobs[0],newjobs[1],newjobs[2])

            

app = QApplication(sys.argv)
w = Exam()

# 윈도우창에 이벤트처리정보를 위젯 객체에 넘겨준다. 
# 메인 루프라고 한다. app.exec_() 끝나야 sys.exit 함수가 호출된다. 
sys.exit(app.exec_())