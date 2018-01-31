#coding:utf-8

try:
    from PyQt4.QtCore import Qt
    from PyQt4.QtGui import QWidget,QLabel,QMessageBox,QApplication,QFont,QGridLayout
    from win32gui import GetCursorPos
    from xlrd import open_workbook
    from sys import exit
    import win32event,pywintypes,win32api,win32gui
    from win32com import client
except:
    exit()
#创建应用程序
try:
    app = QApplication([])
    clipboard = QApplication.clipboard()
    w=QWidget()
except:
    button=QMessageBox.warning(w,"Error",u'创建应用程序出错了！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        exit()

#防止多开
ERROR_ALREADY_EXISTS=183
sz_mutex="test_mutex"
hmutex=win32event.CreateMutex(None,pywintypes.FALSE,sz_mutex)
if (win32api.GetLastError()==ERROR_ALREADY_EXISTS):
    win32api.CloseHandle(hmutex)
    button=QMessageBox.warning(w,"Alert",u'请勿重复开启！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        exit()
        
#创建窗口
try:
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle(u'By ddd')
    #窗口是否为置顶窗口
    #w.setWindowFlags(Qt.FramelessWindowHint|Qt.WindowStaysOnTopHint)
    w.setWindowFlags(Qt.FramelessWindowHint)
    
    label_k=QLabel(u'您输入的关键词:')
    label_r=QLabel(u'匹配的最佳结果:')
    label_fi=QLabel(u'近五年影响因子:')
    label_i=QLabel(u'最新的影响因子:')
    l_keywords=QLabel(u'')
    l_results=QLabel(u'')
    l_if=QLabel(u'')
    l_fiveif=QLabel(u'')
except:
    button=QMessageBox.warning(w,"Error",u'创建界面出错了！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        win32api.CloseHandle(hmutex)
        exit()
        
#用户计算机里没有相关字体库则不设置字体
try:
    label_k.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    label_r.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    label_i.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    label_fi.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    l_keywords.setFont(QFont((u"Roman times"),15,QFont.Bold))
    l_results.setFont(QFont((u"Roman times"),15,QFont.Bold))
    l_if.setFont(QFont((u"Roman times"),15,QFont.Bold))
    l_fiveif.setFont(QFont((u"Roman times"),15,QFont.Bold))
except:
    pass

#控件布局
try:
    grid=QGridLayout()
    grid.addWidget(label_k,1,0)
    grid.addWidget(label_r,2,0)
    grid.addWidget(label_fi,3,0)
    grid.addWidget(label_i,4,0)
    grid.addWidget(l_keywords,1,1)
    grid.addWidget(l_results,2,1)
    grid.addWidget(l_fiveif,3,1)
    grid.addWidget(l_if,4,1)
    w.setLayout(grid)
except:
    button=QMessageBox.warning(w,"Error",u'创建界面出错了！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        win32api.CloseHandle(hmutex)
        exit()

#打开数据表
try:
    table=open_workbook("ddd.xlsx")
    sheet=table.sheets()[0]
    nrows=sheet.nrows
except:
    button=QMessageBox.warning(w,"Error",u'没有检测到数据表或打开数据表失败，请检查数据表ddd.xlsx或使用在线版！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        win32api.CloseHandle(hmutex)
        exit()
    
w.show()

#监听剪贴板内容改变
def on_clipboard_change():
    try:
        data = clipboard.mimeData()
    except:
        button=QMessageBox.warning(w,"Error",u'无法获取剪贴板内容！',QMessageBox.Ok)
        if button==QMessageBox.Ok:
            sys.exit()
            
    if data.hasText():
        #如果剪贴板输入了中文则不处理
        for ch in data.text():
            if ch>=u'\u4e00' and ch<=u'\u9fff':
                return
            
        rawkeywords=str(data.text())
        #如果剪贴板输入了大于120个字符则不处理
        if len(rawkeywords)>120:
            return
        
        global l_keywords,l_results,l_if
        keywords,results,IF,fiveif='','','',''
        try:
        #查询
            #处理换行符，解决在一些pdf文档中可能出现的问题
            keywords=rawkeywords.strip().replace('.','').replace('\r\n',' ').replace('\n',' ').replace('  ',' ')
            keywords2=rawkeywords.strip().replace('.','').replace('\r\n','').replace('\n','').replace('  ',' ')
            for i in range(1,nrows):

                arr=sheet.row_values(i)
                name=str(arr[1])
                shortname=str(arr[0])
                if (name.upper()==keywords.upper()) or (shortname.upper()==keywords.upper()) or (shortname.upper()==keywords2.upper()) or (name.upper()==keywords2.upper()):
                    IF=str(arr[2])
                    fiveif=str(arr[3])
                    results=name
                    break
        except:
            button=QMessageBox.warning(w,"Error",u'无法正确查找，请尝试重启软件！',QMessageBox.Ok)
            if button==QMessageBox.Ok:
                w.close()
                win32api.CloseHandle(hmutex)
                exit()
        
        #显示结果
        
        try:
            
            l_results.setText(results.encode('UTF-8').decode('UTF-8'))
            l_fiveif.setText(fiveif.encode('UTF-8').decode('UTF-8'))
            l_if.setText(IF.encode('UTF-8').decode('UTF-8'))

            l_keywords.setText(str(keywords).encode('UTF-8').decode('UTF-8'))
        
            posX,posY=GetCursorPos()
            w.move(posX,posY)
            w.adjustSize()
            w.update()
            w.show()
            #w.raise_()
            #w.activateWindow()
            shell = client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            window=win32gui.FindWindow(None,u'By ddd')
            win32gui.SetForegroundWindow(window)
            
        except:
            button=QMessageBox.warning(w,"Error",u'无法完成窗口更新，请尝试重启软件！',QMessageBox.Ok)
            if button==QMessageBox.Ok:
                w.close()
                win32api.CloseHandle(hmutex)
                exit()
        
try:        
    clipboard.dataChanged.connect(on_clipboard_change)
except:
    w.close()
    win32api.CloseHandle(hmutex)
    exit()
app.exec_()
