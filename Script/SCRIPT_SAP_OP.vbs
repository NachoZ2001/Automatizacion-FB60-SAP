If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00015"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "F00052"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = "010924"
session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = "010924"
session.findById("wnd[0]/usr/txtBKPF-BKTXT").text = "PAGO ELECTRONCA GARIONE"
session.findById("wnd[0]/usr/txtRF05A-AUGTX").text = "PAGO ELECTRONCA GARIONE"
session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "PAGO ELECTRONCA GARIONE"
session.findById("wnd[0]/usr/ctxtRF05A-AGKON").setFocus
session.findById("wnd[0]/usr/ctxtRF05A-AGKON").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]").close
session.findById("wnd[0]/usr/ctxtRF05A-KONTO").setFocus
session.findById("wnd[0]/usr/ctxtRF05A-KONTO").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB007/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = "*CAJA*"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB007/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = "140000"
session.findById("wnd[0]/usr/ctxtRF05A-AGKON").setFocus
session.findById("wnd[0]/usr/ctxtRF05A-AGKON").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").text = "*ELECTRONI*"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").setFocus
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").close
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").text = "*GARIO*"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").setFocus
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 4
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").setFocus
session.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").caretPosition = 53
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[7]").press
session.findById("wnd[0]/usr/txtBSEG-DMBE2").text = ""
session.findById("wnd[0]/usr/txtBSEG-DMBE3").text = ""
session.findById("wnd[0]/usr/ctxtBSEG-RSTGR").text = "A11"
session.findById("wnd[0]/usr/ctxtBSEG-RSTGR").setFocus
session.findById("wnd[0]/usr/ctxtBSEG-RSTGR").caretPosition = 3
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,6]").setFocus
session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 2
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 0
