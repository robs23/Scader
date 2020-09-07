VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getWeekR 
   Caption         =   "Wybierz tydzień"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "getWeekR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getWeekR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private w As Long
Private y As Long

Private Sub btnOk_Click()
Dim dday As Integer
Dim hhour As Integer
Dim dFrom As Date
Dim dTo As Date
Dim recKeeper As clsRecordsKeeper

w = Me.cmbWeek
y = Me.cmbYear

hhour = 14

Set recKeeper = New clsRecordsKeeper

If Me.optWeek = True Then
    dday = WeekDayName2Int("Niedziela")
    dFrom = DateAdd("h", hhour, Week2Date(CLng(Me.cmbWeek), CLng(Me.cmbYear), dday, vbFirstFourDays))
    dTo = DateAdd("h", 167, dFrom)
Else
    dFrom = DateAdd("h", 6, Me.dFrom)
    dTo = DateAdd("h", 22, Me.dTo)
End If

recKeeper.update dFrom, dTo
'recKeeper.PrintRecords
If cboxTotal.Value = 0 Then
    recKeeper.display False
Else
    recKeeper.display True
End If
Me.Hide

End Sub

Private Sub optDates_Click()
Me.cmbWeek.Enabled = False
Me.cmbYear.Enabled = False
Me.dFrom.Enabled = True
Me.dTo.Enabled = True
End Sub

Private Sub optWeek_Click()
Me.cmbWeek.Enabled = True
Me.cmbYear.Enabled = True
Me.dFrom.Enabled = False
Me.dTo.Enabled = False
End Sub

Private Sub UserForm_Initialize()
For i = Me.cmbWeek.ListCount To 1 Step -1
    Me.cmbWeek.RemoveItem i
Next i

For i = 1 To 53
    Me.cmbWeek.AddItem i
Next i

For i = Me.cmbYear.ListCount To 1 Step -1
    Me.cmbYear.RemoveItem i
Next i

For i = 1 To 10
    Me.cmbYear.AddItem i + 2015
Next i

w = IsoWeekNumber(Date) + 1
y = Year(Date)

Me.dFrom = Date
Me.dTo = DateAdd("d", 7, Date)

Me.optWeek = 1
Me.cmbWeek = w
Me.cmbYear = y

End Sub


Private Sub fillMissing(dFrom As Date, dTo As Date)
Dim i As Integer
Dim sht As Worksheet
Dim curDate As Date
Dim nDate As Date
Dim shift As Integer

Set sht = ThisWorkbook.Sheets("Kontrola upałów")

For i = 2 To 1000
    shift = sht.Cells(i, 3)
    
    If shift = 0 Then
        'we've reached end of table
        Exit For
    End If
Next i

'check if previous line contain shift=1 and curDate=dTo,
'if no, we'll have to add empty lines
If i > 2 Then 'is there any line at all?
    'there's at least 1 line with date
    i = i - 1 'go to previous line
    shift = sht.Cells(i, 3)
    curDate = sht.Cells(i, 1)
Else
    'no data
    shift = 2
    curDate = dFrom
End If

dTo = DateAdd("h", 14, dTo) 'make begining of 2nd shift on next Sunday as the end
Select Case shift
    Case 1
    curDate = DateAdd("h", 6 + 8, curDate) 'first shift
    Case 2
    curDate = DateAdd("h", 14 + 8, curDate) 'second shift
    Case 3
    curDate = DateAdd("h", 22 + 8, curDate) 'third shift
End Select

Do Until curDate >= dTo
     i = i + 1
     Select Case Hour(curDate)
        Case 6: shift = 1
        Case 14: shift = 2
        Case 22: shift = 3
     End Select
     sht.Cells(i, 1) = DateSerial(Year(curDate), Month(curDate), Day(curDate))
     sht.Cells(i, 2) = StrConv(WeekdayName(weekday(curDate, vbSunday), , vbSunday), vbProperCase)
     sht.Cells(i, 3) = shift
     sht.Cells(i, 4) = 0
     curDate = DateAdd("h", 8, curDate)
Loop

End Sub


