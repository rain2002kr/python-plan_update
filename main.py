import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit, QHBoxLayout, QVBoxLayout, QTableWidgetItem, QTableWidget, QGridLayout, QProgressBar
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
        self.btn_timeload = QPushButton('timeload', self)
        self.btn_load = QPushButton('load', self)
        self.btn_save = QPushButton('save', self)
        self.ed_date = QLineEdit('220702', self)
        self.bar1 = QProgressBar(self)
        self.bar1.setOrientation(Qt.Horizontal)

        self.hbox1 = QHBoxLayout()
        self.hbox1.addWidget(self.btn_timeload)
        self.hbox1.addWidget(self.btn_load)
        self.hbox1.addWidget(self.btn_save)
        self.hbox1.addWidget(self.ed_date)
        self.hbox1.addWidget(self.bar1)

        
    def functionCode(self):
        self.btn_timeload.clicked.connect(lambda x: self.main_process('time_load'))
        self.btn_load.clicked.connect(lambda x: self.main_process('load'+ self.ed_date.text()))
        self.btn_save.clicked.connect(lambda x: self.main_process('save'+ self.ed_date.text()))

    def initUI(self):
        self.tbw = {}
        for i in range(0,3):
            self.tbw[i] = QTableWidget()    

        # layout = QVBoxLayout()
        layout = QGridLayout()
        layout.setSpacing(10)
        layout.addLayout(self.hbox1 ,0,0)
        layout.addWidget(self.tbw[0],2,0)
        layout.addWidget(self.tbw[1],3,0)
        layout.addWidget(self.tbw[2],4,0)
        
        self.setLayout(layout)
        
        self.setWindowTitle('하루 기록 저장프로그램')
        self.setGeometry(0,0,500,500)
        self.show()

    def main_process(self, text):
        command = text[0:4]
        date = text[4:]
        print(command)
        print(date)
        
        if text == 'time_load':
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
            for i in range(0,2):
                sumjob = pd.concat([jobs[i],newjobs[i]])
                sumjobs.append(sumjob)
            print(sumjobs)
            # save_time_excel(sumjobs)




            

app = QApplication(sys.argv)
w = Exam()

# 윈도우창에 이벤트처리정보를 위젯 객체에 넘겨준다. 
# 메인 루프라고 한다. app.exec_() 끝나야 sys.exit 함수가 호출된다. 
sys.exit(app.exec_())