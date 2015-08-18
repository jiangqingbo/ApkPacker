#----------------------------------------------------------------------
# A very simple wxPython example.  Just a wx.Frame, wx.Panel,
# wx.StaticText, wx.Button, and a wx.BoxSizer, but it shows the basic
# structure of any wxPython application.
#----------------------------------------------------------------------
# -*- coding:utf-8 -*-  
import wx
import os
import xlrd
import sys
import codecs
import json
import datetime


class MyFrame(wx.Frame):
    """
    This is MyFrame.  It just shows a few controls on a wxPanel,
    and has a simple menu.
    """
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, -1, title,
                          pos=(150, 150), size=(640, 480))

        # Create the menubar
        menuBar = wx.MenuBar()

        # and a menu 
        menu = wx.Menu()
        
        menu.Append(wx.ID_EDIT, "Reset Apks", "clear the settings of apk")

        # add an item to the menu, using \tKeyName automatically
        # creates an accelerator, the third param is some help text
        # that will show up in the statusbar
        menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Exit this simple sample")

        # bind the menu event to an event handler
        self.Bind(wx.EVT_MENU, self.OnTimeToClose, id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.OnResetApks, id=wx.ID_EDIT)

        # and put the menu on the menubar
        menuBar.Append(menu, "&File")
        self.SetMenuBar(menuBar)

        self.CreateStatusBar()
        self.SetStatusText("Powered by Sam, 2014/12", 0)
        

        self.cfgPath = os.getcwd()+"\\config.json"
        cfg = self.loadCofig()
        print cfg
        print os.getcwd()
        
        # Now create the Panel to put the other controls on.
        panel = wx.Panel(self)


        # Use a sizer to layout the controls, stacked vertically and with
        # a 10 pixel border around each
        line1 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenExcel = wx.Button(panel, -1, "Excel")
        self.Bind(wx.EVT_BUTTON, self.OnSelectExcel, btnOpenExcel)
        line1.Add(btnOpenExcel)
        excel = wx.StaticText(panel, -1, "No selected")
        excel.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        excel.SetSize(excel.GetBestSize())
        if cfg != None and cfg.get("excel", None) != None:
            excel.SetLabelText(cfg.get("excel")) 
        line1.Add(excel, 0, wx.ALL,10)
        self.excel = excel
        
        apks = []
        if cfg != None and cfg.get("apk", None) != None:
            apks = cfg.get("apk")
            
        line2 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenApk = wx.Button(panel, -1, "APK(1)")
        self.Bind(wx.EVT_BUTTON, self.OnSelectAPK1, btnOpenApk)
        line2.Add(btnOpenApk)
        apk = wx.StaticText(panel, -1, "No selected")
        apk.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        apk.SetSize(apk.GetBestSize())
        if len(apks) > 0 and apks[0] != None and len(apks[0]) > 1 :
            apk.SetLabelText(apks[0]) 
        line2.Add(apk, 0, wx.ALL,10)
        self.apk1 = apk
        
        line8 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenApk2 = wx.Button(panel, -1, "APK(2)")
        self.Bind(wx.EVT_BUTTON, self.OnSelectAPK2, btnOpenApk2)
        line8.Add(btnOpenApk2)
        apk = wx.StaticText(panel, -1, "No selected")
        apk.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        apk.SetSize(apk.GetBestSize())
        if len(apks) > 1 and apks[1] != None and len(apks[1]) > 1 :
            apk.SetLabelText(apks[1]) 
        line8.Add(apk, 0, wx.ALL,10)
        self.apk2 = apk
        
        line3 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenRar = wx.Button(panel, -1, "WinRar")
        self.Bind(wx.EVT_BUTTON, self.OnSelectWinRar, btnOpenRar)
        line3.Add(btnOpenRar)
        rar = wx.StaticText(panel, -1, "No selected")
        rar.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        rar.SetSize(rar.GetBestSize())
        if cfg != None and cfg.get("rar", None) != None:
            rar.SetLabelText(cfg.get("rar")) 
        line3.Add(rar, 0, wx.ALL,10)
        self.rar = rar
        
        line4 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenSigner = wx.Button(panel, -1, "JarSigner")
        self.Bind(wx.EVT_BUTTON, self.OnSelectJarSigner, btnOpenSigner)
        line4.Add(btnOpenSigner)
        signer = wx.StaticText(panel, -1, "No selected")
        signer.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        signer.SetSize(signer.GetBestSize())
        if cfg != None and cfg.get("signer", None) != None:
            signer.SetLabelText(cfg.get("signer")) 
        line4.Add(signer, 0, wx.ALL,10)
        self.signer = signer
        
        line7 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenKey = wx.Button(panel, -1, "KeyStore")
        self.Bind(wx.EVT_BUTTON, self.OnSelectKeyStore, btnOpenKey)
        line7.Add(btnOpenKey)
        store = wx.StaticText(panel, -1, "No selected")
        store.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        store.SetSize(store.GetBestSize())
        if cfg != None and cfg.get("keystore", None) != None:
            store.SetLabelText(cfg.get("keystore")) 
        line7.Add(store, 0, wx.ALL,10)
        self.store = store
        
        line5 = wx.BoxSizer(wx.HORIZONTAL);
        btnOpenSpace = wx.Button(panel, -1, "WorkSpace")
        self.Bind(wx.EVT_BUTTON, self.OnSelectWorkSpace, btnOpenSpace)
        line5.Add(btnOpenSpace)
        space = wx.StaticText(panel, -1, "No selected")
        space.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        space.SetSize(space.GetBestSize())
        if cfg != None and cfg.get("workspace", None) != None:
            space.SetLabelText(cfg.get("workspace")) 
        line5.Add(space, 0, wx.ALL,10)
        self.space = space

        l1 = wx.Button(panel, -1, "Build Code")
        t1 = wx.TextCtrl(panel, -1, "", size=(300, -1))
        #t1.SetHelpText()
        t1.SetHint("Input build number here, such as 20141209")
        now = datetime.datetime.now()
        ts = now.strftime('%Y%m%d')
        t1.SetLabelText(ts)
        wx.CallAfter(t1.SetInsertionPoint, 0)
        self.buildNumber = t1

        #self.Bind(wx.EVT_TEXT, self.EvtText, t1)
        #t1.Bind(wx.EVT_CHAR, self.EvtChar)
        # t1.Bind(wx.EVT_SET_FOCUS, self.OnSetFocus)
        # t1.Bind(wx.EVT_KILL_FOCUS, self.OnKillFocus)
        # t1.Bind(wx.EVT_WINDOW_DESTROY, self.OnWindowDestroy)

        l2 = wx.Button(panel, -1, "Password")
        t2 = wx.TextCtrl(panel, -1, "", size=(300, -1), style=wx.TE_PASSWORD)
        #self.Bind(wx.EVT_TEXT, self.EvtText, t2)
        self.password = t2

        space = 6

        line6 = wx.FlexGridSizer(cols=3, hgap=space, vgap=space)
        line6.AddMany([ l1, t1, (0,0),
                        l2, t2, (0,0)
                        ])

        #self.Bind(wx.EVT_TEXT, self.EvtText, t1)
        #t1.Bind(wx.EVT_CHAR, self.EvtChar)
        
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.AddSizer(line1)
        sizer.AddSizer(line2)
        sizer.AddSizer(line8)
        sizer.AddSizer(line3)
        sizer.AddSizer(line4)
        sizer.AddSizer(line7)
        sizer.AddSizer(line5)
        sizer.AddSizer(line6)
        btnBuildScript = wx.Button(panel, -1, "Build Dos Script")
        self.Bind(wx.EVT_BUTTON, self.OnBuildScript, btnBuildScript)
        sizer.Add(btnBuildScript, 0, wx.ALL, 10)
        panel.SetSizer(sizer)
        panel.Layout()

        # And also use a sizer to manage the size of the panel such
        # that it fills the frame
        sizer = wx.BoxSizer()
        sizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(sizer)
        

    def OnTimeToClose(self, evt):
        """Event handler for the button click."""
        print "See ya later!"
        self.Close()
        
    def OnResetApks(self, evt):
        """Event handler for the button click."""
        print "OnResetApks"
        self.apk1.SetLabelText("No selected")
        self.apk2.SetLabelText("No selected") 

    def OnFunButton(self, evt):
        """Event handler for the button click."""
        print "Having fun yet?"
    def OnSelectExcel(self, evt):
        """Event handler for the button click."""
        print "OnSelectExcel?"
        wildcard = "Excel file (*.xlsx)|*.xlsx|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose pack configure file",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()

            print paths
            self.excel.SetLabelText(paths[0])
            self.tryExcel(paths[0])

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
    def OnSelectAPK1(self, evt):
        """Event handler for the button click."""
        print "OnSelectAPK"
        wildcard = "Android application (*.apk)|*.apk|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose android application",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()

            self.apk1.SetLabelText(paths[0])

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
    def OnSelectAPK2(self, evt):
        """Event handler for the button click."""
        print "OnSelectAPK"
        wildcard = "Android application (*.apk)|*.apk|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose android application",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()

            self.apk2.SetLabelText(paths[0])

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
    def OnSelectJarSigner(self, evt):
        """Event handler for the button click."""
        print "OnSelectJarSigner"
        wildcard = "application (*.exe)|*.exe|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose jarsigner.exe",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()
            signer = paths[0]
            #jarsigner.exe
            signer = signer[-13:]
            if signer == "jarsigner.exe":
                self.signer.SetLabelText(paths[0])
        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
        
    def OnSelectKeyStore(self, evt):
        """Event handler for the button click."""
        print "OnSelectKeyStore"
        wildcard = "Signature key store (*.keystore)|*.keystore|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose signature key store file",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()
            signer = paths[0]
            self.store.SetLabelText(paths[0])
        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
        
    def OnSelectWinRar(self, evt):
        """Event handler for the button click."""
        print "OnSelectWinRar"
        wildcard = "WinRar application (*.exe)|*.exe|"     \
           "All files (*.*)|*.*"
        # Create the dialog. In this case the current directory is forced as the starting
        # directory for the dialog, and no default file name is forced. This can easilly
        # be changed in your program. This is an 'open' dialog, and allows multitple
        # file selections as well.
        #
        # Finally, if the directory is changed in the process of getting files, this
        # dialog is set up to change the current working directory to the path chosen.
        dlg = wx.FileDialog(
            self, message="Choose winrar application",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard=wildcard,
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            paths = dlg.GetPaths()
            rar = paths[0]
            #WinRAR.exe
            rar = rar[-10:]
            if rar == "WinRAR.exe":
                self.rar.SetLabelText(paths[0])
            

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
    def OnSelectWorkSpace(self, evt):
        """Event handler for the button click."""
        print "OnSelectWorkSpace"
        # In this case we include a "New directory" button. 
        dlg = wx.DirDialog(self, "Choose workspace:",
                          style=wx.DD_DEFAULT_STYLE
                          | wx.DD_DIR_MUST_EXIST
                           #| wx.DD_CHANGE_DIR
                           )

        # If the user selects OK, then we process the dialog's data.
        # This is done by getting the path data from the dialog - BEFORE
        # we destroy it. 
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPath()
            self.space.SetLabelText(paths)

        # Only destroy a dialog after you're done with it.
        dlg.Destroy()
        
    def OnBuildScript(self, evt):
        """Event handler for the button click."""
        print "OnBuildScript"
        if self.excel.GetLabelText() == "No selected":
            dlg = wx.MessageDialog(self, 'Please set excel configure file first',
                               'A Message Box',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
            return
#         if self.apk1.GetLabelText() == "No selected":
#             dlg = wx.MessageDialog(self, 'Please set source apk(1) first',
#                                'A Message Box',
#                                wx.OK | wx.ICON_INFORMATION
#                                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
#                                )
#             dlg.ShowModal()
#             dlg.Destroy()
#             return
#         
#         if self.apk2.GetLabelText() == "No selected":
#             dlg = wx.MessageDialog(self, 'Please set source apk(2) first',
#                                'A Message Box',
#                                wx.OK | wx.ICON_INFORMATION
#                                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
#                                )
#             dlg.ShowModal()
#             dlg.Destroy()
#             return
        if self.rar.GetLabelText() == "No selected":
            dlg = wx.MessageDialog(self, 'Please set WinRar first',
                               'A Message Box',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
            return
        if self.store.GetLabelText() == "No selected":
            dlg = wx.MessageDialog(self, 'Please set signature file first',
                               'A Message Box',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
            return
        if self.signer.GetLabelText() == "No selected":
            dlg = wx.MessageDialog(self, 'Please set JarSigner first',
                               'A Message Box',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
            return
        if self.space.GetLabelText() == "No selected":
            dlg = wx.MessageDialog(self, 'Please set WorkSpace first',
                               'A Message Box',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
            return
        #save configure and build script
        worksapce = self.space.GetLabelText()
        excel = self.excel.GetLabelText()
        rar = self.rar.GetLabelText()
        signer = self.signer.GetLabelText();
        keystore = self.store.GetLabelText();
        apk1 = self.apk1.GetLabelText()
        if apk1 == "No selected":
            apk1 = None
        apk2 = self.apk2.GetLabelText()
        if apk2 == "No selected":
            apk2 = None
        if apk1 == None and apk2 == None: 
            dlg = wx.MessageDialog(self, 'Please set source apk(1) or apk(2) first',
                                'A Message Box',
                                wx.OK | wx.ICON_INFORMATION
                                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                )
            dlg.ShowModal()
            dlg.Destroy()
            return            
        apk = (apk1, apk2)
        self.saveConfig(worksapce, excel, apk, rar, signer, keystore)
        self.buildDosScript(worksapce, excel, apk, rar, signer, self.buildNumber.GetLineText(0), self.password.GetLineText(0), keystore)
        
        
    def tryExcel(self, name):
        book = xlrd.open_workbook(name)
        print "The number of worksheets is", book.nsheets
        print "Worksheet name(s):", book.sheet_names()
        sh = book.sheet_by_index(0)
        print sh.name, sh.nrows, sh.ncols
        paramKeys = []
        if sh.ncols > 7:
            for i in range(7, sh.ncols):
                key = sh.cell_value(rowx=2, colx=i)
                key = key.strip()
                paramKeys.append(key)
        print paramKeys
                    
                
        if sh.nrows > 3 :
            for i in range(3, sh.nrows):
                chanel = sh.cell_value(rowx=i, colx=0)
                name = sh.cell_value(rowx=i, colx=1)
                name =  name.strip()
                if sh.ncols > 7:
                    targetApk = sh.cell_value(rowx=i, colx=6)
                    targetApk = targetApk.strip()
                    if len(targetApk) > 0 :
                        targetApk = "\\"+targetApk
                    print targetApk
                    for j in range(7, sh.ncols):
                        key = paramKeys[j-7]
                        if len(key) > 0:
                            param = sh.cell_value(rowx=i, colx=j)
                            if isinstance(param, float):
                                param = '%d'%param
                            line = "%s=%s"%(key,param)
                            print line + ">>.\\assets\\conf.db\r\n"
                key = key.strip()
                paramKeys.append(key)
                
        book.release_resources() 
        return
    def buildDosScript(self, workspace, excel, apk, rar, signer, code, key, store):
        print 'buildDosScript'
        signCmd = u'"%s" -sigalg SHA1withRSA -digestalg SHA1 -verbose -storepass %s -keypass %s -keystore "%s" -signedjar eStockPre.apk app.apk estock\r\n' % (signer, key, key, store)
        bat = codecs.open("\\".join([workspace, "build.bat"]), 'w', "gbk")
        s1 = "REM global settings\r\nmkdir .\\assets\r\nmkdir .\\out\r\ncopy \""
        s2 = "\" .\\app1.apk\r\n"
        if apk[0] != None:
            bat.write(s1+apk[0]+s2)
            bat.write(u'"%s" d ".\\app1.apk" META-INF\\*\r\n' % rar)
        s3 = "copy \""
        s4 = "\" .\\app2.apk\r\n"
        if apk[1] != None :
            bat.write(s3+apk[1]+s4)
            bat.write(u'"%s" d ".\\app2.apk" META-INF\\*\r\n' % rar)
        #write line by line in excel
        book = xlrd.open_workbook(excel)
        print "The number of worksheets is", book.nsheets
        print "Worksheet name(s):", book.sheet_names()
        sh = book.sheet_by_index(0)
        print sh.name, sh.nrows, sh.ncols
        paramKeys = []
        if sh.ncols > 8:
            for i in range(8, sh.ncols):
                key = sh.cell_value(rowx=2, colx=i)
                key = key.strip()
                paramKeys.append(key)
        
                
        if sh.nrows > 3 :
            for i in range(3, sh.nrows):
                mode = sh.cell_value(rowx=i, colx=7)
                try:
                    if len(mode)>0:
                        mode = mode.lower();
                except Exception as e:
                    mode = ""
                sourceApk = None
                if mode == 'a' and apk[0] != None:
                    sourceApk = "app1.apk"
                if mode == 'b' and apk[1] != None:
                    sourceApk = "app2.apk"
                if sourceApk == None :
                    continue
                chanel = sh.cell_value(rowx=i, colx=0)
                name = sh.cell_value(rowx=i, colx=1)
                name =  name.strip()
                path = u'.\\out\\%s\\%s(%d)' % (code, name,chanel)
                rem = u"\r\n\r\nREM pack "+path+u"\r\n"
                bat.write(rem)
                #out put script
                s = u"echo %d=%s>.\\assets\\conf.db\r\n" % (chanel,code)
                bat.write(s)
                targetApk = ""
                if sh.ncols > 8:
                    targetApk = sh.cell_value(rowx=i, colx=6)
                    targetApk = targetApk.strip()
                    if len(targetApk) > 0 :
                        targetApk = "\\"+targetApk
                    for j in range(8, sh.ncols):
                        key = paramKeys[j-8]
                        if len(key) > 0:
                            param = sh.cell_value(rowx=i, colx=j)
                            if isinstance(param, float):
                                param = '%d'%param
                            s = "echo %s=%s >> .\\assets\\conf.db\r\n" % (key,param)
                            bat.write(s)
                bat.write(u'"%s" A ".\\%s" .\\assets\\conf.db\r\n' % (rar, sourceApk));
                bat.write(signCmd.replace("app.apk", sourceApk))
                bat.write(u"mkdir \""+path+u"\"\r\n")
                bat.write(u"copy eStockPre.apk \""+path+targetApk+u"\"\r\n")
                
        book.release_resources()    
        bat.close()
        hint = 'Please run build.bat under %s'%workspace
        dlg = wx.MessageDialog(self, hint,
                               'Done',
                               wx.OK | wx.ICON_INFORMATION
                               #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                               )
        dlg.ShowModal()
        dlg.Destroy()
        return
    def saveConfig(self, workspace, excel, apk, rar, signer, keystore):
        print 'saveConfig'
        print self.cfgPath
        cfg = codecs.open(self.cfgPath, 'w', "utf-8")
        obj = {'workspace':workspace,'excel':excel, "apk":apk, 'rar':rar, "signer":signer, "keystore":keystore}
        encodedjson = json.dumps(obj)
        cfg.write(encodedjson)
        cfg.close()
        self.loadCofig()
        
    def loadCofig(self):
        print 'loadCofig'
        print self.cfgPath
        try:
            cfg = codecs.open(self.cfgPath, 'r', "utf-8")
            if cfg != None:
                data = json.load(cfg)
            else:
                data = None
        except Exception as e:
            data = None
        print data
        return data


class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, "Apk Packer v0.92")
        self.SetTopWindow(frame)

        print "Print statements go to this stdout window by default."

        frame.Show(True)
        return True

#reload(sys)                         # 2
#sys.setdefaultencoding('gb2312')        
app = MyApp(redirect=False)
app.MainLoop()

