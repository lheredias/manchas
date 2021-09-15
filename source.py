from PyQt5.QtCore import (pyqtSignal,QThreadPool,pyqtSlot,QRunnable,QObject,Qt)
from PyQt5.QtWidgets import (QApplication, QMainWindow,QLabel,QFileDialog,QAction,
                             QProgressBar, QPushButton,QMessageBox,QLineEdit,QMenu,
                             QGraphicsOpacityEffect,QComboBox,QHBoxLayout,QStackedLayout,
                             QTextEdit,QCheckBox,QVBoxLayout,QWidget,QListView)

from PyQt5.QtGui import (QIcon,QFont,QPixmap,QCursor)
from webbrowser import open as op
import sys
import os
import getpass
import requests, json
import pandas as pd
# import numpy as np

from datetime import date
today=date.today()

if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)

if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)
instrucciones=resource_path('INSTRUCCIONES DE USO.pdf')
pic=resource_path('check small.png')

icon=resource_path('finalicon02.ico')
logo=resource_path('app name02.png')

credentials=resource_path('credentials.txt')
with open(credentials) as json_file:
    data = json.load(json_file)
client=data['client_id']
key=data['client_secret']

username=getpass.getuser()
defaultDir='D:\\Usuarios\\'+username+'\\Documents'

choices=['1. Consulta masiva de validez de Comprobantes de Pago'] 

if os.name=='nt':
    fontOne = QFont("Helvetica", 9)
    fontTwo=QFont("Helvetica", 9)
    fontThree=QFont('Consolas', 11)  #Done message font
    fontFive=QFont('Consolas', 11) #Version font
    window_size=[700,540]
    window_move=[0,0]
else:
    fontOne = QFont("San Francisco", 12)
    fontTwo=QFont("San Francisco", 12)
    fontThree=QFont('San Francisco', 13)  #Done message font
    fontFive=QFont('San Francisco', 13) #Version font
    window_size=[900,600]
    window_move=[100,100]

# <codecell>  
class WorkerSignalsNine(QObject):
    alert=pyqtSignal(str)
    finished=pyqtSignal(str)
    report=pyqtSignal(str)
    
class JobRunnerNine(QRunnable):    
    signals = WorkerSignalsNine()
    
    def __init__(self,ruc,client_id,client_secret,SaveAs,origin):
        super().__init__()

        self.is_killed = False 
        self.ruc=ruc
        self.client_id=client_id
        self.client_secret=client_secret
        self.SaveAs=SaveAs
        self.origin=origin
        
    @pyqtSlot()
    
        
    def is_opened(self):
        temp_filename=self.SaveAs[:-4]+' temp.xlsx'
        if os.path.exists(self.SaveAs) == True:
            try:              
                os.rename(self.SaveAs,temp_filename)
                os.rename(temp_filename,self.SaveAs)               
                return False
            except PermissionError:
                return True
        else:
            return False
    
    def translate_estadoCp(self,x):
        if x=='0':
            return 'NO EXISTE'
        elif x=='1':
            return 'ACEPTADO'
        elif x=='2':
            return 'ANULADO'
        elif x=='3':
            return 'AUTORIZADO'
        elif x=='4':
            return 'NO AUTORIZADO'
    def translate_estadoRuc(self,x):
        if x=='00':
            return 'ACTIVO'
        elif x=='01':
            return 'BAJA PROVISIONAL'
        elif x=='02':
            return 'BAJA PROV. POR OFICIO'
        elif x=='03':
            return 'SUSPENSION TEMPORAL'
        elif x=='10':
            return 'BAJA DEFINITIVA'
        elif x=='11':
            return 'BAJA DE OFICIO'
        elif x=='22':
            return 'INHABILITADO-VENT.UNICA'
    def translate_condDomiRuc(self,x):
        if x=='00':
            return 'HABIDO'
        elif x=='09':
            return 'PENDIENTE'
        elif x=='11':
            return 'POR VERIFICAR'
        elif x=='12':
            return 'NO HABIDO'
        elif x=='20':
            return 'NO HALLADO'
    def to_str(self,x):
        if type(x)==list:
            holder=''
            for i in range(len(x)):
                if i==(len(x)-1):
                    holder+='['+x[i].replace('-','').strip()+']'
                else:
                    holder+='['+x[i].replace('-','').strip()+']'+'\n'
                return holder
        else:
            pass
    def get_token(self):
        try:
            grant_type = "client_credentials"
            scope="https://api.sunat.gob.pe/v1/contribuyente/contribuyentes"
            headers = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36"}
            data = {
                "grant_type": grant_type,
                "scope": scope,
                "client_id": self.client_id,
                "client_secret": self.client_secret
                }
            url="https://api-seguridad.sunat.gob.pe/v1/clientesextranet/"+self.client_id+"/oauth2/token/"
            auth_response = requests.post(url, data=data, headers=headers)
            auth_response_json = auth_response.json()
            auth_token = auth_response_json["access_token"]
            token_type=auth_response_json["token_type"]
            expires_in=auth_response_json["expires_in"]
            auth_token_header_value = "Bearer %s" % auth_token
            auth_token_header = {"Authorization": auth_token_header_value}
            return auth_token_header
        except Exception as e:
            self.signals.alert.emit('Error1')       
    def connect(self,rc,output,count):
        new_count=count
        auth_token_header=self.get_token()
        url02 = "https://api.sunat.gob.pe/v1/contribuyente/contribuyentes/"+self.ruc+"/validarcomprobante"
        headers = {'Authorization': auth_token_header}
        for i in range(count,len(rc)):
            if self.is_killed:
                break
            else:
                r = requests.post(url02, json=rc[i], headers=auth_token_header)
                r=r.text
                r=json.loads(r)
                self.signals.report.emit('Procesando '+str(new_count+1)+' de '+str(len(rc))+' CdP: '+r['message']) 
                if r['message']=='Unauthorized':
                    self.signals.report.emit('Generando nuevo Token...')
                    self.connect(self,rc,output,new_count)
                new_count+=1
                output.append(r)
    def get_data(self):
        try:
            rc=self.origin
            rc=pd.read_excel(rc,sheet_name= 'consulta')
            rc['CTIPDOCCOM']=rc['CTIPDOCCOM'].apply(lambda x: '0'+str(x) if len(str(x))<2 else str(x))
            # rc['MONTO']=round(rc['CIMPTOTCOM']/rc['CTIPCAM'],2)
            # rc['MONTO']=rc['MONTO'].apply(lambda x: round(x,2))
            rc['CFECCOM'] = rc['CFECCOM'].apply(lambda x: x if type(x)==str else x.strftime('%d/%m/%Y'))
            safe=rc.copy()
            # rc=rc.drop(columns=['CIMPTOTCOM','CTIPCAM'])
            for i in range(len(rc)):
                if rc['CNUMSER'][i].startswith(('E','F','e','f')):
                    rc['MONTO'][i]=round(rc['MONTO'][i],2)
                else:
                    rc['MONTO'][i]=''
            rc.columns=['fechaEmision','codComp','numeroSerie','numero','numRuc','monto']
            rc=rc[['numRuc','codComp','numeroSerie','numero','fechaEmision','monto']]
            rc=rc.to_json(orient="records")
            rc = json.loads(rc)
            return rc,safe
        except Exception as e:
            self.signals.alert.emit('Error2')
    def run(self):
        
        try:
            output=[]
            rc,safe=self.get_data()
            self.connect(rc,output,count=0)
            if self.is_killed:
                pass
            else:
                df=pd.json_normalize(output)
                df['data.estadoCp']=df['data.estadoCp'].apply(self.translate_estadoCp)
                df['data.estadoRuc']=df['data.estadoRuc'].apply(self.translate_estadoRuc)
                df['data.condDomiRuc']=df['data.condDomiRuc'].apply(self.translate_condDomiRuc)
                if 'data.observaciones' in df.columns:
                    df['data.observaciones']=df['data.observaciones'].apply(self.to_str)
                if 'errorCode' in df.columns:
                    df=df.drop(['errorCode'],axis=1)
                if 'status' in df.columns:
                    df=df.drop(['status'],axis=1)
                df=df.drop(['success'],axis=1)
                
                try:
                    safe=safe.join(df)
                    safe['CFECCOM']=pd.to_datetime(safe['CFECCOM'],format='%d/%m/%Y')
                    with pd.ExcelWriter(self.SaveAs, 
                                            engine='xlsxwriter',
                                            datetime_format='d/mm/yyyy') as writer: 
                             safe.to_excel(writer,sheet_name='resultado',index = False) 
                             workbook  = writer.book
                             format1 = workbook.add_format({'num_format': 'd/mm/yyyy'})
                             worksheet = writer.sheets['resultado']
                             worksheet.set_column('A:A',15, format1)
                    # safe.to_excel(self.SaveAs,index=False)
                except Exception as e:
                    self.signals.alert.emit('Error3') 
                self.signals.report.emit('Proceso completado satisfactoriamente.')
                self.signals.finished.emit('Done')

            #============================================END============================================#   
                            
        except Exception as e:
            self.signals.alert.emit(str(e))                           
               # self.signals.alert.emit(str(type(e)))                        
                
    def kill(self):
        self.is_killed = True
           
class ActionsNine(QWidget):

 
    def __init__(self):
        super().__init__()
        self.runner=None
        self.title = 'Manchas'
        self.var1=None
        self.var2=None
        self.var4=None
        self.var7=None
        self.initUI()
        self.msg1='Verifica los datos ingresados.' 
        self.msg2='Esta versión ya caducó.' 
        
    def initUI(self):
        self.style = QApplication.style()
       
        
        self.style1=("QPushButton { background-color: rgb(155, 61, 61 ); color: rgb(255, 255, 255 );}")
        self.style2=("QPushButton { background-color: rgb(69, 70, 77); color: rgb(255, 255, 255);}") 
        self.style3 = ("QProgressBar {border: 2px solid grey;border-radius: 5px;text-align: center}"
                         "QProgressBar::chunk {background-color: IndianRed;width: 10px;margin: 1px;}") 
        
        self.setWindowTitle(self.title)
        
        
        self.v1=QVBoxLayout()
        self.h1=QHBoxLayout()
        self.h2=QHBoxLayout()
        self.h4=QHBoxLayout()
        self.h5=QHBoxLayout()
        self.h6=QHBoxLayout()
        self.h7=QHBoxLayout()
        self.h8=QHBoxLayout()
        self.h9=QHBoxLayout()
        self.v2=QVBoxLayout()
                               
        #self.setStyleSheet("background-color: rgb(255, 255, 255); color: rgb(86, 88, 110)")
        
        self.setWindowIcon(QIcon(icon))
    
    
        self.buttonOne = QPushButton('Id', self)
        self.buttonOne.setMinimumHeight(35)
        # self.buttonOne.setMaximumWidth(200)
        self.buttonOne.setStyleSheet(self.style2)
        self.buttonOne.setFont(fontTwo)
        self.buttonOne.setEnabled(False)
        self.buttonOne.setCursor(QCursor(Qt.PointingHandCursor))
        self.h1.addWidget(self.buttonOne,1)
        
        self.buttonTwo = QPushButton('Clave', self)
        self.buttonTwo.setMinimumHeight(35)
        # self.buttonTwo.setMaximumWidth(200)
        self.buttonTwo.setStyleSheet(self.style2)
        self.buttonTwo.setFont(fontTwo)
        self.buttonTwo.setEnabled(False)
        self.buttonTwo.setCursor(QCursor(Qt.PointingHandCursor)) 
        self.h2.addWidget(self.buttonTwo,1)
        
        self.buttonFour = QPushButton('RUC', self)
        self.buttonFour.setMinimumHeight(35)
        # self.buttonFour.setMaximumWidth(200)
        self.buttonFour.setStyleSheet(self.style2)
        self.buttonFour.setFont(fontTwo)
        self.buttonFour.setEnabled(False)
        self.buttonFour.setCursor(QCursor(Qt.PointingHandCursor))
        self.h4.addWidget(self.buttonFour,1)
        
        self.buttonFive = QPushButton('Cargar Excel', self)      
        self.buttonFive.clicked.connect(self.openFileNameDialogOne)
        self.buttonFive.setMinimumHeight(35)
        self.buttonFive.setStyleSheet(self.style2)
        self.buttonFive.setFont(fontTwo)
        self.buttonFive.setCursor(QCursor(Qt.PointingHandCursor))
        self.h6.addWidget(self.buttonFive,1)
        
        self.buttonSix = QPushButton('Guardar como', self)      
        self.buttonSix.clicked.connect(self.openFileNameDialogTwo)
        self.buttonSix.setMinimumHeight(35)
        self.buttonSix.setStyleSheet(self.style2)
        self.buttonSix.setFont(fontTwo)
        self.buttonSix.setCursor(QCursor(Qt.PointingHandCursor))
        self.h8.addWidget(self.buttonSix,1)
        
        self.myTextBoxOne = QLineEdit(self)      
        # self.myTextBoxOne.setEchoMode(QLineEdit.Password)
        self.myTextBoxOne.setMinimumHeight(35)  
        self.myTextBoxOne.setStyleSheet('background-color: rgb(69, 70, 77); color: white')
        self.myTextBoxOne.setFont(fontTwo)
        self.myTextBoxOne.setText(client)
        self.h1.addWidget(self.myTextBoxOne,4)
        
        self.myTextBoxTwo = QLineEdit(self)
        # self.myTextBoxTwo.setEchoMode(QLineEdit.Password)
        self.myTextBoxTwo.setMinimumHeight(35)  
        self.myTextBoxTwo.setStyleSheet('background-color: rgb(69, 70, 77); color: white')  
        self.myTextBoxTwo.setFont(fontTwo)
        self.myTextBoxTwo.setText(key)
        self.h2.addWidget(self.myTextBoxTwo,4)
         
        self.myTextBoxFour = QLineEdit(self)
        self.myTextBoxFour.setMinimumHeight(35)  
        self.myTextBoxFour.setStyleSheet('background-color: rgb(69, 70, 77); color: white')  
        self.myTextBoxFour.setFont(fontTwo)
        self.myTextBoxFour.setPlaceholderText('Número de RUC del receptor')
        self.h4.addWidget(self.myTextBoxFour,4)
        
        self.myTextBoxFive = QLineEdit(self)
        self.myTextBoxFive.setMinimumHeight(35)  
        self.myTextBoxFive.setStyleSheet('background-color: rgb(69, 70, 77); color: white')  
        self.myTextBoxFive.setFont(fontTwo)
        self.myTextBoxFive.setPlaceholderText('Elige el destino')      
        self.myTextBoxFive.setReadOnly(True)
        self.h8.addWidget(self.myTextBoxFive,4)
        
        self.myTextBoxThree = QLineEdit(self)
        self.myTextBoxThree.setMinimumHeight(35)  
        self.myTextBoxThree.setStyleSheet('background-color: rgb(69, 70, 77); color: white')  
        self.myTextBoxThree.setFont(fontTwo)
        self.myTextBoxThree.setPlaceholderText('Carga tu archivo Excel')      
        self.myTextBoxThree.setReadOnly(True)
        self.h7.addWidget(self.myTextBoxThree,2)
        
        self.h9.addStretch()                   
        self.start = QPushButton('Ejecutar', self)
        self.start.setStyleSheet(self.style1)
        # self.start.setFocus()
        self.start.setFont(fontOne)
        self.start.setMinimumHeight(35)
        # self.start.setMaximumWidth(200)
        self.start.setEnabled(True)
        self.start.setCursor(QCursor(Qt.PointingHandCursor))
        self.start.clicked.connect(self.started) 
        self.h9.addWidget(self.start)
    
        self.button = QPushButton('Limpiar', self)
        self.button.setStyleSheet(self.style1)
        self.button.setFont(fontOne)
        self.button.setMinimumHeight(35)
        # self.button.setMaximumWidth(200)
        self.button.setEnabled(True)
        self.button.setCursor(QCursor(Qt.PointingHandCursor))
        self.button.clicked.connect(self.clean) 
        self.h9.addWidget(self.button)
        
        self.progress = QProgressBar(self)
        self.progress.setFormat("")
        self.progress.setStyleSheet(self.style3)    
        self.progress.setFont(fontOne)
        # self.progress.setMaximumWidth(800)
        self.progress.setAlignment(Qt.AlignCenter) 
        self.progress.setValue(0)
        self.progress.setMaximum(0)
        self.progress.hide()
        
        self.labelOne = QLabel('', self)
        self.labelOne.setFont(fontThree)
        self.labelOne.setAlignment(Qt.AlignCenter)
        self.labelOne.hide()
        
        self.labelTwo = QLabel('', self)
        self.labelTwo.setFont(fontThree)
        self.labelTwo.setStyleSheet("color:LightGreen")
        self.labelTwo.setAlignment(Qt.AlignCenter)
        # self.labelTwo.hide()      

        self.effect = QGraphicsOpacityEffect(self)
        self.pixmap = QPixmap(pic)
        self.pixmap = self.pixmap.scaled(50, 50, Qt.KeepAspectRatio,Qt.SmoothTransformation)
        self.labelThree = QLabel('', self)
        self.labelThree.setAlignment(Qt.AlignCenter)
       
        self.report=QTextEdit(self)   
        self.report.setFont(fontTwo)
        self.report.setPlaceholderText('Acá se generará el reporte del proceso...') 
        # self.report.setText('Acá se generará el reporte del proceso...') 
        self.report.setStyleSheet("color: Gainsboro;border: 2px solid rgb(69, 70, 77)")     
        self.report.setReadOnly(True)
        self.v2.addWidget(self.report)
        
        self.mainLayout = QHBoxLayout()
        self.mainLayout.setAlignment(Qt.AlignCenter)
        self.v1.setAlignment(Qt.AlignCenter)
        # self.mainLayout.setSpacing(30)
        self.v1.addLayout(self.h1)
        self.v1.addLayout(self.h2)
        self.v1.addLayout(self.h4)
        self.h5.addLayout(self.h6,1)
        self.h5.addLayout(self.h7,4)
        self.v1.addLayout(self.h5)
        self.v1.addLayout(self.h8)       
        self.v1.addLayout(self.h9) 
        self.v1.addWidget(self.progress)
        self.v1.addWidget(self.labelOne)
        self.v1.addWidget(self.labelTwo)
        self.v1.addWidget(self.labelThree)    
        self.mainLayout.addLayout(self.v1)
        self.mainLayout.addLayout(self.v2)
        self.setLayout(self.mainLayout)
        
        # quit = QAction("Quit", self)
        # quit.triggered.connect(self.closeEvent)
   
    def started(self):
        
        if today!=date(2021, 12, 25):

            if self.runner is None:
                self.start.setEnabled(False)
                self.var1=self.myTextBoxOne.text().strip()
                self.var2=self.myTextBoxTwo.text().strip()
                self.var4=self.myTextBoxFour.text().strip()
                if (self.var1 and self.var2 and self.var4 and self.var7) is not None and len(self.var4)==11:
                    cred={'client_id':self.var1,'client_secret':self.var2}
                    with open(credentials, 'w') as outfile:
                        json.dump(cred, outfile)
                    self.button.setEnabled(False)
                    self.labelTwo.setText('')
                    self.labelThree.hide()
                    self.progress.show()
                    self.threadpool = QThreadPool()
                    self.runner = JobRunnerNine(self.var4,self.var1,self.var2,self.var7,self.var6)   
                    self.threadpool.start(self.runner)                                         
                    try:
                        self.runner.signals.alert.disconnect(self.alert)
                        self.runner.signals.finished.disconnect(self.finished)
                        self.runner.signals.report.disconnect(self.report_msg)
                    except TypeError:     
                        self.runner.signals.alert.connect(self.alert)
                        self.runner.signals.finished.connect(self.finished)
                        self.runner.signals.report.connect(self.report_msg)
                    else:
                        self.runner.signals.alert.connect(self.alert)
                        self.runner.signals.finished.connect(self.finished)
                        self.runner.signals.report.connect(self.report_msg)
                else:
                    self.start.setEnabled(True)
                    self.labelTwo.setText('Intenta de nuevo.')
                    self.error(self.msg1)
                    self.progress.hide()
        else:
            self.start.setEnabled(True)
            self.labelTwo.setText('Actualiza la aplicación.')
            self.error(self.msg2)
            self.progress.hide()   
    def clean(self):

        self.myTextBoxThree.setText(None)
        self.myTextBoxFour.setText(None)
        self.myTextBoxFive.setText(None)
        self.var1=None
        self.var2=None
        self.var4=None
        self.var6=None
        self.var7=None
        self.runner=None
        self.labelTwo.setText('')
        self.labelOne.setText('')
        self.labelThree.hide()
        self.report.setText(None)
        self.progress.hide()

    def openFileNameDialogOne(self):
        
        fileName, _ = QFileDialog.getOpenFileName(self,"Selecciona tu documento",'',filter="PDF (*.xlsx)")
        
        if fileName:        
            if '.xlsx' not in fileName:
                fileName=fileName+'.xlsx'
            fileName=os.path.abspath(fileName)         
            self.myTextBoxThree.setText(fileName)
            self.var6=self.myTextBoxThree.text()
        return fileName 
    def openFileNameDialogTwo(self):
        
        fileName, _ = QFileDialog.getSaveFileName(self,"Guardar como",'',filter="Excel (*.xlsx)")
        
        if fileName:        
            if '.xlsx' not in fileName:
                fileName=fileName+'.xlsx'
            fileName=os.path.abspath(fileName)         
            self.myTextBoxFive.setText(fileName)
            self.var7=self.myTextBoxFive.text()
        return fileName
    
    def alert(self, msg):
        if msg=='Error1':
            self.error('No se pudo generar el Token. Asegúrate de haber ingresado el "Id" y "Clave" correctos. Si el problema persiste, puede que el servidor de SUNAT presente problemas. Intenta de nueva más tarde.')
        elif msg=='Error2':
            self.error('Ocurrió un problema al leer el archivo Excel. Asegúrate de que la estructura sea la indicada.')
        elif msg=='Error3':
            self.error('El archivo Excel sobre el cual intentas guardar el resultado se encuentra abierto.')
        elif msg=='Error4':
            self.error('Esta versión ya caducó. Actualíza la aplicación.')
        else:
            self.error('Ocurrió un error inesperado: '+msg)
        self.clean()
    def report_msg(self,msg):
        self.report.append(msg)
    def finished(self, msg):
        if msg=='Done':
            self.runner=None
            self.myTextBoxThree.setText(None)
            self.myTextBoxFour.setText(None)
            self.myTextBoxFive.setText(None)
            self.var1=None
            self.var2=None
            self.var4=None
            self.start.setEnabled(True)   
            self.labelTwo.setText('¡Listo, ya puedes visualizar tus documentos!')
            self.labelThree.show()
            self.labelThree.setPixmap(self.pixmap) 
            self.progress.hide()
            self.button.setEnabled(True) 
    
    def error(self,errorMsg):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(self.title)
        msg.setWindowIcon(QIcon(icon))
        msg.setText("Error")
        msg.setFont(fontTwo)
        msg.setStandardButtons(QMessageBox.Ok)
        buttonOk = msg.button(QMessageBox.Ok)
        buttonOk.setCursor(QCursor(Qt.PointingHandCursor))
        buttonOk.setFont(fontOne)
        msg.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(69, 70, 77  )")
        msg.setInformativeText(errorMsg)
        msg.exec_()
        self.start.setEnabled(True)
        self.button.setEnabled(True)
        self.runner=None

    def instructions(self):
        if os.name=='nt':
            os.startfile(instrucciones)
        else:
            import subprocess
            subprocess.run(['open', instrucciones], check=True)    


     
# <codecell>  
    
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.window1 = ActionsNine()
        self.title = 'Manchas'
        self.initUI()
        
    def initUI(self):  

        self.style1=("QPushButton { background-color: rgb(155, 61, 61 ); color: rgb(255, 255, 255 );}"
                     "QPushButton:hover { background-color: rgba(155, 61, 61,230) ;color: white;}"
                      "QPushButton:pressed { background-color: rgb(69, 70, 77) ;color: rgb(255, 255, 255 );}")
        self.style2=("QPushButton { background-color: rgb(69, 70, 77); color: rgb(255, 255, 255);}"
                      "QPushButton:hover { background-color: rgba(69, 70, 77,230) ;color: white;}"
                      "QPushButton:pressed { background-color: rgb(155, 61, 61 ); color: rgb(255, 255, 255 );}")
        self.style4=("QComboBox {selection-background-color: rgb(69, 70, 77);background-color: rgb(69, 70, 77); color: rgb(255, 255, 255);padding-left:10px}"
                     "QComboBox QAbstractItemView::item { min-height: 35px; min-width: 50px;}"
                     "QListView::item { color: white; background-color: rgb(69, 70, 77)}"
                     "QListView::item:selected { color: white; background-color: IndianRed}") 
        
        self.style = QApplication.style()
       
        self.setWindowTitle(self.title)       
        # self.setMinimumSize(750,500)
        self.setMinimumSize(window_size[0],window_size[1])
        # self.resize(500,600)
        self.move(window_move[0],window_move[1])
        # self.setWindowState(Qt.WindowMaximized)
        self.setStyleSheet("background-color: rgb(22, 23, 24); color:CornflowerBlue")
        self.setWindowIcon(QIcon(icon))
        
        self.menuBar = self.menuBar()
        self.menuBar.setCursor(QCursor(Qt.PointingHandCursor))
        self.menuBar.setStyleSheet("QMenuBar {background-color: rgb(155, 61, 61); color: rgb(255, 255, 255)}"
                                   "QMenuBar:item:selected {background-color: white ;color: black}") 
        
        self.menu=QMenu("&Menú")
        self.menu.setStyleSheet("QMenu {background-color: white; color: black}"
                                   "QMenu:item:selected {background-color: white ;color: rgb(155, 61, 61)}") 
        self.menuBar.addMenu(self.menu)
        self.menu.addAction('&Acerca de', self.about)
        self.menu.addAction('&Ir al repositorio', self.repo)
        self.menu.addAction('&Actualizar', self.update)
        self.menu.addAction('&Instrucciones de uso', self.window1.instructions)

        self.stackedLayout = QStackedLayout()
              
        self.mainLayout = QVBoxLayout()
        self.mainLayout.setAlignment(Qt.AlignCenter)    
    
        self.h=QHBoxLayout()
        self.v=QVBoxLayout()
        
        self.v0=QVBoxLayout()
        self.v1=QVBoxLayout()
        self.v2=QVBoxLayout()
        self.v3=QVBoxLayout()
        self.h3=QHBoxLayout()
        self.h4=QHBoxLayout()       
        self.h5=QHBoxLayout()
        
        windows=[self.window1]    
        
        for window in windows:
            self.stackedLayout.addWidget(window)
            
        self.pageCombo = QComboBox()   
        self.pageCombo.addItems(choices)
        self.pageCombo.setMinimumHeight(35)
        self.pageCombo.setStyleSheet(self.style4)
        self.listview=QListView()
        self.listview.setFont(fontTwo)
        self.listview.setCursor(QCursor(Qt.PointingHandCursor))
        self.pageCombo.setView(self.listview)
        self.pageCombo.setCursor(QCursor(Qt.PointingHandCursor))
        self.pageCombo.setFont(fontTwo)
        self.pageCombo.activated.connect(self.toggle_window)

        self.v0.addWidget(self.pageCombo)

        self.h.addLayout(self.v1)
        self.h.addLayout(self.v)
        self.h.addLayout(self.v2)   

        self.stackedLayout.setAlignment(Qt.AlignCenter)
        self.h.setAlignment(Qt.AlignCenter)
               
        self.mainLayout.addLayout(self.h,1)   
        self.mainLayout.addLayout(self.v0,0)   
        
        self.mainLayout.addLayout(self.stackedLayout,4)        
      
        self.pixmap = QPixmap(icon)
        self.pixmap = self.pixmap.scaled(75, 75, Qt.KeepAspectRatio,Qt.SmoothTransformation)
        self.labelThree = QLabel('', self)
        self.labelThree.setPixmap(self.pixmap) 
        self.labelThree.setAlignment(Qt.AlignCenter) 
        self.v1.addWidget(self.labelThree)
        
        self.logo = QPixmap(logo)
        self.logo = self.logo.scaled(130, 130, Qt.KeepAspectRatio,Qt.SmoothTransformation)
        self.labelFour = QLabel('', self)
        self.labelFour.setPixmap(self.logo) 
        self.labelFour.setAlignment(Qt.AlignCenter) 
        self.v.addWidget(self.labelFour)
        
        self.titleOne = QLabel('Versión 1.0', self)
        self.titleOne.setFont(fontFive)
        self.titleOne.setStyleSheet("color:	IndianRed")
        self.titleOne.setAlignment(Qt.AlignRight | Qt.AlignBottom)  
        self.v2.addWidget(self.titleOne)
        
        self.labelOne = QLabel('Hola, '+username, self)
        self.labelOne.setFont(fontFive)
        self.labelOne.setAlignment(Qt.AlignRight)  
        self.v2.addWidget(self.labelOne)        
        
        self.status_label = QLabel()
        self.statusBar().addPermanentWidget(self.status_label)
        self.status_label.setText('Versión 1.0 lanzada en septiembre del 2021.')

        self.w = QWidget(self)
        self.w.setLayout(self.mainLayout)
        self.setCentralWidget(self.w)
        
        quit = QAction("Quit", self)
        quit.triggered.connect(self.closeEvent)
        
    def toggle_window(self):
        self.stackedLayout.setCurrentIndex(self.pageCombo.currentIndex())
    def error(self,errorMsg):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(self.title)
        msg.setWindowIcon(QIcon(icon))
        msg.setText("Error")
        msg.setFont(fontTwo)
        msg.setStandardButtons(QMessageBox.Ok)
        buttonOk = msg.button(QMessageBox.Ok)
        buttonOk.setCursor(QCursor(Qt.PointingHandCursor))
        buttonOk.setFont(fontOne)
        msg.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(69, 70, 77  )")
        msg.setInformativeText(errorMsg)
        msg.exec_()

    def closeEvent(self, event):
        close = QMessageBox()
        close.setWindowTitle("¿Estás seguro?")
        close.setWindowIcon(QIcon(icon))
        close.setFont(fontTwo)
        close.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(69, 70, 77  )")
        if self.window1.runner: 
            close.setText("La aplicación está consolidando el detalle de FE recibidas. Si sales, se detendrá por completo el proceso y no se guardarán los resultados.")
            # QMessageBox.information(self,'Finalizando...','El proceso se detendrá por completo en unos instantes.')
        else:
            close.setText("Se abandonará por completo la aplicación.")           
        close.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        buttonYes = close.button(QMessageBox.Yes)
        buttonYes.setCursor(QCursor(Qt.PointingHandCursor))
        buttonYes.setFont(fontOne)
        buttonYes.setText('Sí')
        buttonCancel = close.button(QMessageBox.Cancel)
        buttonCancel.setText('No')
        buttonCancel.setCursor(QCursor(Qt.PointingHandCursor))
        buttonCancel.setFont(fontOne)
        close = close.exec()

        if close == QMessageBox.Yes:  
            if self.window1.runner: 
                self.window1.runner.kill()               
            event.accept() 
        else:
            event.ignore()
    def repo(self):
        op('https://github.com/lheredias/manchas')
    def about(self):
        info = QMessageBox()
        info.setWindowTitle("Acerca de Manchas")
        
        info.setWindowIcon(QIcon(icon))
        info.setText('''Manchas permite hacer la consulta masiva de Comprobantes de pago a través de una interfaz gráfica de usuario.''')

        info.setFont(fontTwo)
        info.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(69, 70, 77  )")
        info.setWindowModality(0)
        # info.setModal(True)
        info.activateWindow()
        info.setStandardButtons(QMessageBox.Ok)
        buttonOk = info.button(QMessageBox.Ok)
        buttonOk.setText('Entendido')
        buttonOk.setCursor(QCursor(Qt.PointingHandCursor))
        buttonOk.setFont(fontOne)
        info.setDefaultButton(QMessageBox.Ok)
        info.show()
        retval = info.exec_()    
       
    def update(self):
        info = QMessageBox()
        info.setWindowTitle("¿Cómo actualizar Manchas?")
        
        info.setWindowIcon(QIcon(icon))
        info.setText('''Descarga la versión más reciente e instálala. La nueva versión reemplazará automáticamente la anterior.''')

        info.setFont(fontTwo)
        info.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(69, 70, 77  )")
        info.setWindowModality(0)
        info.activateWindow()
        info.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        buttonYes = info.button(QMessageBox.Yes)
        buttonYes.setCursor(QCursor(Qt.PointingHandCursor))
        buttonYes.setText('Buscar')
        buttonYes.setFont(fontOne)
        buttonCancel = info.button(QMessageBox.Cancel)
        buttonCancel.setCursor(QCursor(Qt.PointingHandCursor))
        buttonCancel.setText('Entendido')
        buttonCancel.setFont(fontOne)
        info.setDefaultButton(QMessageBox.Cancel)
        info.show()
        retval = info.exec_()
        print(retval)
        if retval==16384:
            op('https://github.com/lheredias/manchas/releases')
    
if __name__ == '__main__':
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setAttribute(Qt.AA_EnableHighDpiScaling,True)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app.setWindowIcon(QIcon(icon))
    w = MainWindow()
    w.show() 
    sys.exit(app.exec_())
    
