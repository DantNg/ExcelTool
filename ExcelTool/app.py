import sys
# pip install pyqt5
from PyQt5.QtWidgets import QApplication, QMainWindow,QFileDialog
from Gui import Ui_MainWindow
import pandas as pd
import threading

class MainWindow:
    def __init__(self):
        self.main_win = QMainWindow()
        self.uic = Ui_MainWindow()
        self.uic.setupUi(self.main_win)
        self.uic.openFileBtn.clicked.connect(self.openFile)
        self.uic.saveFileBtn.clicked.connect(self.exportFile)
        self.fname=''
        self.foldername=''
        self.DF1 = pd.DataFrame()
        self.DF2 = pd.DataFrame()
        self.checkReadFileComplete = False
      
        #Khởi tạo giá trị mặc định
        self.brand = 'ZTT'
        self.low_vol_dis_default = 43.2
        self.low_vol_dis = self.low_vol_dis_default
        self.low_vol_thres = 47.5
        self.low_vol_thres_default = 49
        self.uic.disconnectionAdjust.setText(str(self.low_vol_dis_default))
        self.uic.threadsholdAdjust.setText(str(self.low_vol_thres))
        self.uic.brandNameAdjust.setText(self.brand)
        self.output_file = 'ket-qua1.xlsx'
        self.uic.nameExportFile.setText(str(self.output_file))

    def show(self):
        # command to run
        self.main_win.show()
    def openFile(self):
        self.uic.saveFileBtn.setEnabled(False)
        self.checkReadFileComplete = False
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.fname, _ = QFileDialog.getOpenFileName(self.main_win,"Select a file:", "","Excel Files (*.xlsx)", options=options)
        if self.fname:
            print(self.fname)
            self.uic.importFilePath.setText(str(self.fname))
            self.uic.statusLabel.setText("Đang tải file ...")
            try :
                t1 = threading.Thread(name='read file', target=self.preProcessSheet)
                t1.start()
            except:
                print("Không thể đọc file!")

    def exportFile(self):
        
        if self.uic.nameExportFile.text() =='' or self.uic.exportFilePath.text() == '':
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            self.foldername = QFileDialog.getExistingDirectory(None, 'Select a folder:', 'C:\\', QFileDialog.ShowDirsOnly)
            if self.foldername:
                print(self.foldername)
                self.uic.exportFilePath.setText(str(self.foldername))
                try:
                    self.uic.statusLabel.setText("Đang xử lí vui lòng chờ ...")
                    t2 = threading.Thread(name='daemon', target=self.processSheet)
                    t2.start()
                    #self.processSheet()
                except:
                    self.uic.statusLabel.setText("Có lỗi vui lòng thử lại!!")             
        elif self.uic.nameExportFile.text() !='' and  self.uic.exportFilePath.text() != '':
                try:
                    self.uic.statusLabel.setText("Đang xử lí vui lòng chờ ...")
                    t2 = threading.Thread(name='daemon', target=self.processSheet)
                    t2.start()
                   
                except:
                    self.uic.statusLabel.setText("Có lỗi vui lòng thử lại!!")
        else:
                 self.uic.statusLabel.setText("Vui lòng nhập thông tin vào ô!")
    def preProcessSheet(self):
        xls = pd.ExcelFile(self.fname)
        try:
            self.DF1 = pd.read_excel(xls,'Accu', usecols="A,D,N,P")
            self.DF2 = pd.read_excel(xls,'Mất điện',usecols="D")
        except:
            print("Lỗi mở file!")
            self.uic.statusLabel.setText("File không hỗ trợ!")
         #Tìm và thay thế các ô trống 
        cond = self.DF1['Tủ nguồn DC'] == ' '
        self.DF1.loc[cond,'Tủ nguồn DC'] = self.DF1['Mã trạm']
        self.DF1.rename(columns = {'Tủ nguồn DC':'Mã tủ nguồn DC'}, inplace = True)
        self.uic.statusLabel.setText("Đọc file xong!")
        self.uic.saveFileBtn.enabled = True
        self.uic.saveFileBtn.setEnabled(True)

    def processSheet(self):
                df1 = self.DF1
                df2 = self.DF2
                #print(df1
                print(self.uic.disconnectionAdjust.text())
                print(self.uic.threadsholdAdjust.text())
                #Đặt mặc định cho cột low voltage disconnection
                if self.uic.disconnectionAdjust.text() != '' :
                    try:
                        df1['Battery low voltage disconnection (Vdc)'] = float(self.uic.disconnectionAdjust.text())
                    except :
                        df1['Battery low voltage disconnection (Vdc)'] = self.low_vol_dis_default
                #Đặt hãng cần chỉnh
                if self.uic.brandNameAdjust.text() != '' :
                    self.cond1 = df1['Hãng sản xuất'] == self.uic.brandNameAdjust.text() 
                else:
                    self.cond1 = df1['Hãng sản xuất'] == self.brand 
                
                #Đặt mặc định cho cột low voltage threadshold
                df1['Battery low voltage threadshold (Vdc)'] = self.low_vol_thres_default
                if self.uic.threadsholdAdjust.text() != '' :
                    try:
                        df1.loc[self.cond1,'Battery low voltage threadshold (Vdc)'] = float(self.uic.threadsholdAdjust.text())
                        #df1['Battery low voltage disconnection (Vdc)'] = float(self.uic.threadsholdAdjust.text())
                    except :
                        df1.loc[self.cond1,'Battery low voltage threadshold (Vdc)'] = self.low_vol_thres

                # Filtering Dataframe rows
                #df1.to_csv('ouput1.csv',index = False)
                if self.uic.exportLostStation.isChecked():
                    df3 = df1[df1[['Mã trạm']].agg(tuple,1).isin(df2[['Mã trạm']].agg(tuple,1))]
                    df3 = df3.drop_duplicates(subset=df1.columns.difference(['STT']))
                    df3['STT'] = range(1, len(df3) + 1) #đánh số lại
                    df3 = df3[['STT','Mã tủ nguồn DC','Battery low voltage disconnection (Vdc)','Battery low voltage threadshold (Vdc)']]
                    print(df3)
                    path = self.foldername+'/'+self.uic.nameExportFile.text()
                    df3.to_excel(path,index = False)
                else:
                    df1 = df1[['STT','Mã tủ nguồn DC','Battery low voltage disconnection (Vdc)','Battery low voltage threadshold (Vdc)']]
                    #lọc hàng trùng
                    df1 = df1.drop_duplicates(subset=df1.columns.difference(['STT']))
                    df1['STT'] = range(1, len(df1) + 1) #đánh số lại
                    print(df1)
                    path = self.foldername+'/'+self.uic.nameExportFile.text()
                    df1.to_excel(path,index = False)
                self.uic.statusLabel.setText("Xuất file thành công!!")
                

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())