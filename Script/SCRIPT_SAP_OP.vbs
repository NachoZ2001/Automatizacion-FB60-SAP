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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "F00053"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00015"

Dim objFSO, objExcel, objWorkbook, objSheet, scriptPath, excelPath, i
Set objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
excelPath = objFSO.BuildPath(scriptPath, "Data\ExcelSap.xlsx")
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(excelPath)
Set objSheet = objWorkbook.Sheets(1)

lastRow = objSheet.Cells(objSheet.Rows.Count, 25).End(-4162).Row
i = 2

Do While i <= lastRow
    importe = Trim(CStr(objSheet.Cells(i, 35).Value))
    fecha_convertida = Trim(CStr(objSheet.Cells(i, 20).Value))
    cuentaMayor = Trim(CStr(objSheet.Cells(i, 21).Value))
    indicadorImpuesto = Trim(CStr(objSheet.Cells(i, 19).Value))
    texto = Trim(CStr(objSheet.Cells(i, 18).Value))
    If Len(texto) > 50 Then
    	texto = Left(texto, 50)
    End If
    acreedor = Trim(CStr(objSheet.Cells(i, 22).Value))
    referencia = Trim(CStr(objSheet.Cells(i, 23).Value))
    centroCosto = Trim(CStr(objSheet.Cells(i, 24).Value))
    fechaContabilidad = Trim(CStr(objSheet.Cells(i, 25).Value))  
    tipoVenta = Trim(CStr(objSheet.Cells(i, 27).Value))
    categoriaVenta = Trim(CStr(objSheet.Cells(i, 28).Value))
    tipoMapeado = Trim(CStr(objSheet.Cells(i, 29).Value))
    fecha_OP = Trim(CStr(objSheet.Cells(i, 31).Value))
    texto_OP = Trim(CStr(objSheet.Cells(i, 32).Value))
    cuenta_OP = Trim(CStr(objSheet.Cells(i, 33).Value))
    tipo_OP = Trim(CStr(objSheet.Cells(i, 34).Value))

    session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = fecha_OP
    session.findById("wnd[0]/usr/txtBKPF-BKTXT").text = texto_OP
    session.findById("wnd[0]/usr/txtRF05A-AUGTX").text = texto_OP
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = texto_OP
    session.findById("wnd[0]/usr/ctxtRF05A-KONTO").text = cuenta_OP
    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importe
    session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = acreedor
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    While statusBarMessageType = "E" Or statusBarMessageType = "W"
    	session.findById("wnd[0]").sendVKey 0
        statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    Wend
    If tipo_OP = "CAJA" Then
	session.findById("wnd[0]/tbar[1]/btn[14]").press
    	statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	While statusBarMessageType = "E" Or statusBarMessageType = "W"
            session.findById("wnd[0]").sendVKey 0
            statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	Wend
    	session.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").setFocus 
    	session.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").caretPosition = 26
    	session.findById("wnd[0]").sendVKey 2
    	session.findById("wnd[0]/tbar[1]/btn[7]").press
    	statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	While statusBarMessageType = "E" Or statusBarMessageType = "W"
    	    session.findById("wnd[0]").sendVKey 0
            statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	Wend
    	session.findById("wnd[0]/usr/txtBSEG-DMBE2").text = ""
    	session.findById("wnd[0]/usr/txtBSEG-DMBE3").text = ""
    	session.findById("wnd[0]/usr/ctxtBSEG-RSTGR").text = "A11"
    	session.findById("wnd[0]/tbar[1]/btn[14]").press
    	statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	While statusBarMessageType = "E" Or statusBarMessageType = "W"
    	    session.findById("wnd[0]").sendVKey 0
            statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    	Wend
    End If

    session.findById("wnd[0]/tbar[0]/btn[11]").press
    statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    While statusBarMessageType = "E" Or statusBarMessageType = "W"
    	session.findById("wnd[0]").sendVKey 0
        statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    Wend
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    While statusBarMessageType = "E" Or statusBarMessageType = "W"
    	session.findById("wnd[0]").sendVKey 0
        statusBarMessageType = session.findById("wnd[0]/sbar").MessageType
    Wend

    i = i + 1
Loop
