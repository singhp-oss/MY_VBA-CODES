Attribute VB_Name = "TAXDATA"
Sub ImportAirlineTaxInvoiceFinal()

    Dim dlg As FileDialog
    Dim srcPath As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim lastSrcRow As Long, lastDstRow As Long
    Dim i As Long, addedRows As Long
    Dim awbVal As String
    Dim pctDone As Double

    ' --- Step 1: File Picker ---
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Airline Tax Invoice Excel File Select Karo"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = False
    End With

    If dlg.Show = False Then
        MsgBox "Koi file select nahi ki. Macro band ho raha hai.", vbExclamation
        Exit Sub
    End If

    srcPath = dlg.SelectedItems(1)

    ' --- Step 2: Source File Open ---
    Application.ScreenUpdating = False
    Application.StatusBar = "File open ho rahi hai..."
    Set srcWb = Workbooks.Open(Filename:=srcPath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)

    lastSrcRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    If lastSrcRow < 2 Then
        MsgBox "Source file mein koi data nahi mila!", vbExclamation
        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- Step 3: Destination Sheet TAX_DATA ---
    On Error GoTo SheetNotFound
    Set dstWs = ThisWorkbook.Sheets("TAX_DATA")
    On Error GoTo 0
    lastDstRow = dstWs.Cells(dstWs.Rows.Count, 1).End(xlUp).Row + 1
    addedRows = 0

    ' --- Step 4: Data Loop ---
    For i = 2 To lastSrcRow
        pctDone = (i - 1) / (lastSrcRow - 1)
        Application.StatusBar = "Data import ho raha hai... " & Format(pctDone, "0%")

        ' --- AWB Number ---
        awbVal = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Left(awbVal, 4) = "312-" Then awbVal = Mid(awbVal, 5)

        dstWs.Cells(lastDstRow, 1).NumberFormat = "@"
        dstWs.Cells(lastDstRow, 1).Value = awbVal
        dstWs.Cells(lastDstRow, 1).HorizontalAlignment = xlCenter

        ' --- Date ---
        If IsDate(srcWs.Cells(i, 3).Value) Then
            dstWs.Cells(lastDstRow, 2).Value = CDate(srcWs.Cells(i, 3).Value)
            dstWs.Cells(lastDstRow, 2).NumberFormat = "dd/mm/yyyy"
            dstWs.Cells(lastDstRow, 2).HorizontalAlignment = xlCenter
        End If

        ' --- Tax Amount ---
        If IsNumeric(srcWs.Cells(i, 21).Value) Then
            dstWs.Cells(lastDstRow, 3).Value = CDbl(srcWs.Cells(i, 21).Value)
            dstWs.Cells(lastDstRow, 3).NumberFormat = "#,##0.00"
            dstWs.Cells(lastDstRow, 3).HorizontalAlignment = xlCenter
        End If

        lastDstRow = lastDstRow + 1
        addedRows = addedRows + 1
    Next i

    ' --- Step 5: Close Source ---
    srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "? Import Successful!" & vbNewLine & _
           addedRows & " rows TAX_DATA mein add ho gaye!", _
           vbInformation, "Import Complete"
    Exit Sub

SheetNotFound:
    MsgBox "? Error: 'TAX_DATA' sheet nahi mili!" & vbNewLine & _
           "Sheet ka naam exactly 'TAX_DATA' hona chahiye.", _
           vbCritical, "Sheet Not Found"
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub

