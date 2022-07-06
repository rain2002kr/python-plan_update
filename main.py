import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLineEdit
from PyQt5.QtCore import QCoreApplication
from write_dailyplan_to_sql import main

class Exam(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        btn = QPushButton('btn1', self)
        btn.resize(btn.sizeHint())
        btn.move(20,30)
        btn.clicked.connect(self.main_process)

        edit = QLineEdit('edit', self)

        edit.move(120,30)
        
        
        self.setGeometry(300,300,400,500)
        self.setWindowTitle('첫 번째 학습')
        self.show()

    def main_process(self):
        main()

app = QApplication(sys.argv)
w = Exam()

# 윈도우창에 이벤트처리정보를 위젯 객체에 넘겨준다. 
# 메인 루프라고 한다. app.exec_() 끝나야 sys.exit 함수가 호출된다. 
sys.exit(app.exec_())