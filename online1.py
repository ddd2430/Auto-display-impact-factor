#coding:utf-8

try:
    from PyQt4.QtCore import *
    from PyQt4.QtGui import *
    import sys
    import win32gui
    import win32event,pywintypes,win32api
    from win32com import client

    from selenium import webdriver
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import NoSuchElementException
    from selenium.common.exceptions import NoAlertPresentException
except:
    sys.exit()

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

#创建pyqt4界面
try: 
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle(u'By ddd')
    w.setWindowFlags(Qt.FramelessWindowHint)

    label_k=QLabel(u'您输入的关键词:')
    label_r=QLabel(u'匹配的最佳结果:')
    label_i=QLabel(u'最新的影响因子:')
    l_keywords=QLabel(u'')
    l_results=QLabel(u'')
    l_if=QLabel(u'')
except:
    button=QMessageBox.warning(w,"Error",u'创建界面出错了！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        win32api.CloseHandle(hmutex)
        sys.exit()
    #button=QPushButton(u'更多...',w)
#用户计算机里没有相关字体库则不设置字体
try:
    #button.setFont(QFont((u"微软雅黑"),13,QFont.Bold))

    label_k.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    label_r.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    label_i.setFont(QFont((u"微软雅黑"),13,QFont.Bold))
    l_keywords.setFont(QFont((u"Roman times"),15,QFont.Bold))
    l_results.setFont(QFont((u"Roman times"),15,QFont.Bold))
    l_if.setFont(QFont((u"Roman times"),15,QFont.Bold))
except:
    pass

try:
    grid=QGridLayout()
    grid.addWidget(label_k,1,0)
    grid.addWidget(label_r,2,0)
    grid.addWidget(label_i,3,0)
    grid.addWidget(l_keywords,1,1)
    grid.addWidget(l_results,2,1)
    grid.addWidget(l_if,3,1)

    #grid.addWidget(button,4,0)
    
    w.setLayout(grid) 
except:
    button=QMessageBox.warning(w,"Error",u'创建界面出错了！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        win32api.CloseHandle(hmutex)
        sys.exit()

#打开浏览器和网站
try:
    driver=webdriver.PhantomJS("phantomjs-2.1.1-windows\\bin\\phantomjs.exe")
    driver.get("http://www.letpub.com.cn/index.php?page=journalapp&view=search")
except:
    button=QMessageBox.warning(w,"Error",u'无法打开数据网站，请检查phantomJS的完整性或您的网络连接！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        w.close()
        win32api.CloseHandle(hmutex)
        sys.exit()
    
#显示窗口界面    
w.show()

#'更多'按钮被点击时，打开网站
def on_button_clicked():
    webbrowser.open(url)

#监听剪贴板内容改变/槽函数
def on_clipboard_change():
    try:
        data = clipboard.mimeData()
    except:
        button=QMessageBox.warning(w,"Error",u'无法获取剪贴板内容！',QMessageBox.Ok)
        if button==QMessageBox.Ok:
            w.close()
            win32api.CloseHandle(hmutex)
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
        
        try:
        #查询
            try:
                #需要处理换行符
                keywords=rawkeywords.strip().replace('.','').replace('\n',' ').replace('\r\n',' ').replace('  ',' ')
            except:
                pass
            IF,results='',''
            driver.find_element_by_id("searchname").clear()
            driver.find_element_by_id("searchname").send_keys(keywords)
            driver.find_element_by_css_selector("input[style=\"width: 86px; background:url(images/yxyz_search.jpg) no-repeat left top; height: 25px; border:0 none; font-weight: bold; cursor:pointer; vertical-align:middle;\"]").click()
            
            try:
                IF=driver.find_element_by_xpath("//table[1]/tbody[1]/tr[3]/td[3]").text
                results=driver.find_element_by_xpath("//table[1]/tbody[1]/tr[3]/td[2]/a[@target]").text
            except:
                IF,results='',''
                   
        except:
            button=QMessageBox.warning(w,"Error",u'未知错误，请尝试重启软件！',QMessageBox.Ok)
            if button==QMessageBox.Ok:
                w.close()
                win32api.CloseHandle(hmutex)
                sys.exit()
        
        #显示结果
        try:
            l_results.setText(str(results).encode('UTF-8').decode('UTF-8'))
            l_if.setText(str(IF).encode('UTF-8').decode('UTF-8'))
            l_keywords.setText(str(keywords).encode('UTF-8').decode('UTF-8'))
            
            posX,posY=win32gui.GetCursorPos()
            w.move(posX,posY)
            w.adjustSize()
            w.update()
            #w.activateWindow()
            w.show()
            shell = client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            window=win32gui.FindWindow(None,u'By ddd')
            win32gui.SetForegroundWindow(window)
            
        except:
            button=QMessageBox.warning(w,"Error",u'更新数据出错！',QMessageBox.Ok)
            if button==QMessageBox.Ok:
                w.close()
                win32api.CloseHandle(hmutex)
                sys.exit()

#给剪贴板绑定槽函数        
try:        
    clipboard.dataChanged.connect(on_clipboard_change)
    #button.clicked.connect(on_button_clicked)
except:
    button=QMessageBox.warning(w,"Error",u'绑定槽函数出错！',QMessageBox.Ok)
    if button==QMessageBox.Ok:
        w.close()
        win32api.CloseHandle(hmutex)
        sys.exit()

app.exec_()
