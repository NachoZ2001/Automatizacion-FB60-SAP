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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00002"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00002"

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
    importe = Trim(CStr(objSheet.Cells(i, 17).Value))
    fecha_convertida = Trim(CStr(objSheet.Cells(i, 20).Value))
    cuentaMayor = Trim(CStr(objSheet.Cells(i, 21).Value))
    indicadorImpuesto = Trim(CStr(objSheet.Cells(i, 19).Value))
    texto = Trim(CStr(objSheet.Cells(i, 18).Value))
    acreedor = Trim(CStr(objSheet.Cells(i, 22).Value))
    referencia = Trim(CStr(objSheet.Cells(i, 23).Value))
    centroCosto = Trim(CStr(objSheet.Cells(i, 24).Value))
    fechaContabilidad = Trim(CStr(objSheet.Cells(i, 25).Value))  
    tipoVenta = Trim(CStr(objSheet.Cells(i, 27).Value))
    categoriaVenta = Trim(CStr(objSheet.Cells(i, 28).Value))
    tipoMapeado = Trim(CStr(objSheet.Cells(i, 29).Value))	

    If categoriaVenta = "GASTOS COMBUSTIBLES" Then
        neto = Trim(CStr(objSheet.Cells(i, 12).Value))
        iva = Trim(CStr(objSheet.Cells(i, 16).Value))
        noGravado = Trim(CStr(objSheet.Cells(i, 13).Value))
        indicadorRenglon1 = "C1"
        indicadorRenglon2 = "C0"
        netoConIva = Trim(CStr(objSheet.Cells(i, 26).Value))
    End If

    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-ACCNT").text = acreedor
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = fecha_convertida
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-XBLNR").text = referencia
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BUDAT").text = fechaContabilidad
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-WRBTR").text = importe
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/chkINVFO-XMWST").setFocus
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/chkINVFO-XMWST").selected = true
    session.findById("wnd[0]").sendVKey 0
    If tipoMapeado != "C" Then
    	session.findById("wnd[0]").sendVKey 0
    End If
    session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").text = texto

    If categoriaVenta = "GASTOS COMBUSTIBLES" Then
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = cuentaMayor
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,1]").text = cuentaMayor
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").text = netoConIva
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,1]").text = noGravado
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-MWSKZ[6,0]").text = indicadorRenglon1
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-MWSKZ[6,1]").text = indicadorRenglon2
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-ZUONR[9,0]").text = Left(texto, 18)
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-ZUONR[9,1]").text = Left(texto, 18)
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,0]").text = texto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,1]").text = texto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[20,0]").text = centroCosto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[20,1]").text = centroCosto
    Else
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").text = cuentaMayor
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").text = importe
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-MWSKZ[6,0]").text = indicadorImpuesto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-ZUONR[9,0]").text = Left(texto, 18)
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,0]").text = texto
        session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[20,0]").text = centroCosto
    End If

    session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[20,0]").setFocus
    session.findById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[20,0]").caretPosition = 10
    If tipoMapeado = "C" Then
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]").sendVKey 0
    End If
    session.findById("wnd[0]/tbar[1]/btn[9]").press
    session.findById("wnd[0]").sendVKey 0
    If tipoMapeado != "C" Then
    	session.findById("wnd[0]").sendVKey 0
    End If
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    objSheet.Cells(i, 30).Value = "Cargado correctamente"

    i = i + 1
Loop

objWorkbook.Close False
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing