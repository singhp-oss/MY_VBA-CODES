Attribute VB_Name = "transferdata"
Sub TransferToMaster()
    Dim MasterPath As String, MasterSheetName As String, MasterFileName As String
    Dim wbMaster As Workbook, wsMaster As Worksheet, wsDaily As Worksheet
    Dim LastRowMaster As Long, LastRowDaily As Long, i As Long
    Dim LastSN As Long
    Dim ThisWB As Workbook
    
    ' --- SETTINGS ---
    MasterPath = "Z:\MASTER_DATABASE.xlsm"
    MasterFileName = "MASTER_DATABASE.xlsm"
    MasterSheetName = "DATA_2025-26"
    
    Set ThisWB = ThisWorkbook
    Set wsDaily = ThisWB.ActiveSheet
    LastRowDaily = wsDaily.Cells(wsDaily.Rows.Count, "B").End(xlUp).Row ' Column B check karega
    
    If LastRowDaily < 2 Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- MASTER FILE OPEN/CHECK ---
    On Error Resume Next
    Set wbMaster = Workbooks(MasterFileName)
    On Error GoTo 0
    
    If wbMaster Is Nothing Then
        If Dir(MasterPath) <> "" Then
            Set wbMaster = Workbooks.Open(MasterPath, UpdateLinks:=0)
        Else
            MsgBox "Master file nahi mili! Z: Drive check karein.", vbCritical
            GoTo CleanExit
        End If
    End If
    
    Set wsMaster = wbMaster.Sheets(MasterSheetName)
    
    ' --- DATA TRANSFER LOOP ---
    Dim dataTransferred As Boolean
    dataTransferred = False
    
    For i = 2 To LastRowDaily
        ' Check agar AF column khali hai aur B column mein data hai
        If wsDaily.Cells(i, 32).Value = "" And wsDaily.Cells(i, 2).Value <> "" Then
            
            ' Master ki sabse aakhri row (Column B ke hisaab se)
            LastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, "B").End(xlUp).Row + 1
            
            ' Serial Number Update (Pichli row ka SN dekh kar +1)
            LastSN = 0
            On Error Resume Next
            LastSN = wsMaster.Cells(LastRowMaster - 1, 1).Value
            On Error GoTo 0
            wsMaster.Cells(LastRowMaster, 1).Value = LastSN + 1
            
            ' B se AE tak ka data transfer
            wsMaster.Range("B" & LastRowMaster & ":AE" & LastRowMaster).Value = _
            wsDaily.Range("B" & i & ":AE" & i).Value
            
            wsDaily.Cells(i, 32).Value = "TRANSFER DONE"
            dataTransferred = True
        End If
    Next i
    
    If dataTransferred Then
        wbMaster.Save
        MsgBox "Data Transfer Ho Gaya!", vbInformation
    Else
        MsgBox "Koi naya data nahi mila.", vbExclamation
    End If

CleanExit:
    ThisWB.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

