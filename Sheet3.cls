VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private m_OldVal As String

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Count = 1 Then
        m_OldVal = Trim(CStr(Target.Value))
    Else
        m_OldVal = ""
    End If

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    '=================================
    ' COMMODITY ALERT (COLUMN AG)
    '=================================
    If Target.Count = 1 And Target.Column = 33 And Target.Row >= 2 Then

        Dim commodityCode As String
        commodityCode = UCase(Trim(CStr(Target.Value)))

        Select Case commodityCode

            Case "XPS", "HEA"

                If Me.Cells(Target.Row, 1).Interior.Color <> RGB(255, 199, 206) Then
                    Me.Range(Me.Cells(Target.Row, 1), Me.Cells(Target.Row, 32)).Interior.Color = RGB(255, 255, 0)
                End If

                MsgBox "SPECIAL COMMODITY ALERT !" & vbCrLf & vbCrLf & _
                       "Commodity Code : " & commodityCode & vbCrLf & _
                       "Row No : " & Target.Row, _
                       vbExclamation, "Commodity Alert"

            Case Else

                If Me.Cells(Target.Row, 1).Interior.Color = RGB(255, 255, 0) Then
                    Me.Range(Me.Cells(Target.Row, 1), Me.Cells(Target.Row, 32)).Interior.pattern = xlNone
                End If

        End Select

    End If


    '=================================
    ' DUPLICATE AWB CHECK (COLUMN B)
    '=================================
    If Target.Count > 1 Then Exit Sub
    If Target.Column <> 2 Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If m_OldVal <> "" Then Exit Sub

    Dim newVal As String
    newVal = LCase(Trim(CStr(Target.Value)))

    If newVal = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = Me.Cells(Me.Rows.Count, 2).End(xlUp).Row

    Dim r As Long
    Dim scanned As Long
    Dim cv As String

    scanned = 0

    For r = lastRow To 1 Step -1

        If r = Target.Row Then GoTo nextRow

        cv = LCase(Trim(CStr(Me.Cells(r, 2).Value)))

        If cv <> "" Then

            scanned = scanned + 1

            If scanned > 50 Then Exit For

            If cv = newVal Then

                Me.Range(Me.Cells(r, 1), Me.Cells(r, 32)).Interior.Color = RGB(255, 199, 206)

                Me.Range(Me.Cells(Target.Row, 1), Me.Cells(Target.Row, 32)).Interior.Color = RGB(255, 199, 206)

                MsgBox "DUPLICATE AWB FOUND!" & vbCrLf & vbCrLf & _
                       "AWB : " & UCase(newVal) & vbCrLf & _
                       "Row : " & r & " mein pehle se exist karta hai." & vbCrLf & vbCrLf & _
                       "Dono rows pink highlight ho gayi hain." & vbCrLf & _
                       "Please verify karein.", _
                       vbExclamation, "Duplicate AWB Alert"

                Exit For

            End If

        End If

nextRow:
    Next r

End Sub



