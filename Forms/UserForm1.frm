VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StartDate As Date
Private endDate As Date



Private Sub btnDo_Click()
StartDate = CDate(Me.dateFrom.Value & " " & Me.timeFrom.Value)
endDate = CDate(Me.dateTo.Value & " " & Me.timeTo.Value)


If StartDate > endDate Then
    MsgBox "Data zakończenia musi być późniejsza od daty startu!", vbOKOnly + vbExclamation, "Niewłaściwy zakres"
Else
    'MsgBox "Data startu: " & startDate & " " & startTime & ", data zakończenia: " & endDate & " " & endTime
    If DateDiff("d", StartDate, endDate) > 120 Then
        MsgBox "Wybrano zbyt duży zakres czasowy. Wybierz maksymalnie 10 dni!", vbOKOnly + vbExclamation, "Zbyt duży zakres"
    Else
        ThisWorkbook.Sheets(1).Cells.Clear
        ThisWorkbook.Sheets("RN3000").Cells.Clear
        ThisWorkbook.Sheets("RN4000").Cells.Clear
        If ThisWorkbook.Sheets(1).ChartObjects.Count > 0 Then
            ThisWorkbook.Sheets(1).ChartObjects.Delete
        End If
        If Me.txtBlends.Value = "" Then
            If Me.txtExclude.Value = "" Then
                If IsNull(Me.cmbPiec.Value) Or Me.cmbPiec.Value = "Wszystkie" Or Me.cmbPiec.Value = "" Then
                    Connection StartDate, endDate
                Else
                    Connection StartDate, endDate, CStr(Right(Me.cmbPiec.Value, 4))
                End If
            Else
                If IsNull(Me.cmbPiec.Value) Or Me.cmbPiec.Value = "Wszystkie" Or Me.cmbPiec.Value = "" Then
                    Connection StartDate, endDate, , , getBlends(UserForm1.txtExclude.Value)
                Else
                    Connection StartDate, endDate, CStr(Right(Me.cmbPiec.Value, 4)), , getBlends(UserForm1.txtExclude.Value)
                End If
            End If
        Else
            If Me.txtExclude.Value = "" Then
                If IsNull(Me.cmbPiec.Value) Or Me.cmbPiec.Value = "Wszystkie" Or Me.cmbPiec.Value = "" Then
                    Connection StartDate, endDate, , getBlends(UserForm1.txtBlends.Value)
                Else
                    Connection StartDate, endDate, CStr(Right(Me.cmbPiec.Value, 4)), getBlends(UserForm1.txtBlends.Value)
                End If
            Else
                If IsNull(Me.cmbPiec.Value) Or Me.cmbPiec.Value = "Wszystkie" Or Me.cmbPiec.Value = "" Then
                    Connection StartDate, endDate, , getBlends(UserForm1.txtBlends.Value), getBlends(UserForm1.txtExclude.Value)
                Else
                    Connection StartDate, endDate, CStr(Right(Me.cmbPiec.Value, 4)), getBlends(UserForm1.txtBlends.Value), getBlends(UserForm1.txtExclude.Value)
                End If
            End If
        End If
        Me.Hide
    End If
End If
End Sub


Private Sub CheckBox1_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub txt3000min_Change()

End Sub

Private Sub UserForm_Initialize()
Me.dateFrom.Value = Date
Me.dateTo.Value = Date
setTimeFrom
setTimeTo
setPiec
End Sub

Sub setTimeFrom()
Dim i As Integer
Dim h As String
Dim m As String

Me.timeFrom.Clear
h = -1
For i = 0 To 48
    If i Mod 2 = 0 Then
        h = h + 1
        m = "00"
    Else
        m = "30"
    End If
    Me.timeFrom.AddItem h & ":" & m
Next i
End Sub

Sub setTimeTo()
Dim i As Integer
Dim h As String
Dim m As String

Me.timeTo.Clear
h = -1
For i = 0 To 48
    If i Mod 2 = 0 Then
        h = h + 1
        m = "00"
    Else
        m = "30"
    End If
    Me.timeTo.AddItem h & ":" & m
Next i
End Sub

Sub setPiec()
Me.cmbPiec.Clear
Me.cmbPiec.AddItem "Wszystkie", 0
Me.cmbPiec.AddItem "RN3000", 1
Me.cmbPiec.AddItem "RN4000", 2
End Sub

Function getBlends(txt As String) As Long()
Dim v() As String
Dim arr() As Long
Dim i As Integer
v = Split(txt, ",", , vbTextCompare)

If Not isArrayEmpty(v) Then
    For i = LBound(v) To UBound(v)
        If IsNumeric(Trim(v(i))) Then
            If isArrayEmpty(arr) Then
                ReDim arr(0) As Long
                arr(0) = CLng(Trim(v(i)))
            Else
                ReDim Preserve arr(UBound(arr) + 1) As Long
                arr(UBound(arr)) = CLng(Trim(v(i)))
            End If
        End If
    Next i
    getBlends = arr
End If

End Function
