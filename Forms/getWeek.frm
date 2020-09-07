VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getWeek 
   Caption         =   "Wybierz tydzień"
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4290
   OleObjectBlob   =   "getWeek.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getWeek"
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

w = Me.cmbWeek
y = Me.cmbYear

hhour = 14

dday = WeekDayName2Int("Niedziela")
dFrom = DateAdd("h", hhour, Week2Date(CLng(Me.cmbWeek), CLng(Me.cmbYear), dday, vbFirstFourDays))
dTo = DateAdd("h", 160, dFrom)

update dFrom, dTo
Me.Hide

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

Me.cmbWeek = w
Me.cmbYear = y

End Sub

Private Sub update(dFrom As Date, dTo As Date)
Dim sql As String
Dim rs As ADODB.Recordset
Dim sht As Worksheet

On Error GoTo err_trap

updateConnection

Set sht = ThisWorkbook.Sheets("Kontrola mielenia")

sht.Range("A2:E30").Cells.ClearContents

sql = "DECLARE @sDate as datetime; " _
    & "DECLARE @eDate as datetime; " _
    & "SELECT @sDate = '" & Format(dFrom, "yyyy-mm-dd hh:mm") & "'; " _
    & "SELECT @eDate='" & Format(dTo, "yyyy-mm-dd hh:mm") & "'; " _
    & "SELECT CONVERT(date,od.plMoment,103) as Data, DATEPART(weekday,CONVERT(date,od.plMoment,103)) as Weekday, CASE WHEN DATEPART(hh,od.plMoment)=6 THEN 1 ELSE CASE WHEN DATEPART(hh,od.plMoment)=14 THEN 2 ELSE 3 END END as [Zmiana], SUM(od.plAmount) as [KG] " _
    & "FROM tbOperations o LEFT JOIN tbOperationData od ON od.operationId=o.operationId " _
    & "WHERE od.plMoment BETWEEN @sDate AND @eDate AND o.type='g' " _
    & "GROUP BY CONVERT(date,od.plMoment,103), DATEPART(weekday,CONVERT(date,od.plMoment,103)), CASE WHEN DATEPART(hh,od.plMoment)=6 THEN 1 ELSE CASE WHEN DATEPART(hh,od.plMoment)=14 THEN 2 ELSE 3 END END " _
    & "ORDER BY Data, Zmiana ASC"
    
Set rs = CreateObject("adodb.recordset")
rs.Open sql, adoConn
Set rs = rs.NextRecordset
Set rs = rs.NextRecordset
If rs.EOF Then
    MsgBox "Brak danych dla wybranego okresu", vbInformation + vbOKOnly, "Brak danych"
Else
    i = 2
    Do Until rs.EOF
        sht.Cells(i, 1) = rs.Fields("Data")
        sht.Cells(i, 2) = StrConv(WeekdayName(rs.Fields("Weekday"), , vbSunday), vbProperCase)
        sht.Cells(i, 3) = rs.Fields("Zmiana")
        sht.Cells(i, 4) = rs.Fields("KG")
        i = i + 1
        rs.MoveNext
    Loop
    fillMissing dFrom, dTo
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""Update"" of getWeek. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub fillMissing(dFrom As Date, dTo As Date)
Dim i As Integer
Dim sht As Worksheet
Dim curDate As Date
Dim nDate As Date
Dim shift As Integer

Set sht = ThisWorkbook.Sheets("Kontrola mielenia")

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
