import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QListWidget, QFileDialog, QMessageBox, QHBoxLayout, QLabel)
from PyQt5.QtCore import Qt
import win32com.client as win32

class HwpxMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('HWPX 파일 병합기')
        self.setGeometry(100, 100, 600, 500)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        title_label = QLabel("병합할 HWPX 파일을 순서대로 추가하세요")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title_label)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.setDragDropMode(QListWidget.InternalMove) # 드래그 앤 드롭으로 순서 변경 가능
        layout.addWidget(self.file_list)
        
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton('파일 추가')
        self.btn_add.clicked.connect(self.addFiles)
        btn_layout.addWidget(self.btn_add)
        
        self.btn_remove = QPushButton('선택 삭제')
        self.btn_remove.clicked.connect(self.removeFiles)
        btn_layout.addWidget(self.btn_remove)
        
        self.btn_clear = QPushButton('전체 삭제')
        self.btn_clear.clicked.connect(self.file_list.clear)
        btn_layout.addWidget(self.btn_clear)
        
        layout.addLayout(btn_layout)
        
        self.btn_merge = QPushButton('파일 병합 및 저장')
        self.btn_merge.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 50px; font-size: 16px;")
        self.btn_merge.clicked.connect(self.mergeFiles)
        layout.addWidget(self.btn_merge)
        
        info_label = QLabel("※ 이 프로그램은 한컴오피스가 설치된 환경에서 작동합니다.")
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(info_label)

    def addFiles(self):
        files, _ = QFileDialog.getOpenFileNames(self, "HWPX 파일 선택", "", "HWPX Files (*.hwpx);;All Files (*)")
        if files:
            self.file_list.addItems(files)

    def removeFiles(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def mergeFiles(self):
        count = self.file_list.count()
        if count < 2:
            QMessageBox.warning(self, "경고", "병합할 파일을 2개 이상 선택해주세요.")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(self, "저장할 파일명 입력", "", "HWPX Files (*.hwpx)")
        if not save_path:
            return
            
        try:
            # 한글 객체 생성
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.XHwpWindows.Item(0).Visible = False # 백그라운드 실행
            
            # 첫 번째 파일 열기
            first_file = self.file_list.item(0).text()
            hwp.Open(first_file)
            
            # 두 번째 파일부터 이어붙이기
            for i in range(1, count):
                file_path = self.file_list.item(i).text()
                hwp.MovePos(3) # 문서 끝으로 이동
                # InsertFile(파일경로, 다음페이지부터삽입여부)
                # "NextPage"는 다음 페이지부터, ""는 현재 커서 위치부터
                hwp.InsertFile(file_path, "NextPage") 
                
            hwp.SaveAs(save_path)
            hwp.Quit()
            
            QMessageBox.information(self, "완료", f"파일 병합이 완료되었습니다.\n저장 위치: {save_path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"병합 중 오류가 발생했습니다: {str(e)}\n(한컴오피스가 설치되어 있는지 확인해주세요)")
            if 'hwp' in locals():
                try:
                    hwp.Quit()
                except:
                    pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = HwpxMergerApp()
    ex.show()
    sys.exit(app.exec_())
