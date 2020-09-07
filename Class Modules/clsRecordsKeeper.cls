VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecordsKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pRecords As New Collection
Private roastingBatches As New Collection
Private pDFrom As Date
Private pDTo As Date

Public Function getRoastingBatches(Optional ind As Variant) As Variant
If Not IsMissing(ind) Then
    Set getRoastingBatches = roastingBatches(ind)
Else
    Set getRoastingBatches = roastingBatches
End If

End Function

Public Sub update(dFrom As Date, dTo As Date)
pDFrom = dFrom
pDTo = dTo
Dim sql As String
Dim rs As ADODB.Recordset
Dim nRec As clsRecord

updateConnection

sql = "DECLARE @sDate as datetime; " _
        & "DECLARE @eDate as datetime; " _
        & "SELECT @sDate = '" & Format(dFrom, "yyyy-mm-dd hh:mm") & "'; " _
        & "SELECT @eDate='" & Format(dTo, "yyyy-mm-dd hh:mm") & "'; " _
        & "SELECT od.plMoment as [Data], z.zfinIndex as [Index], m.machineName as [Maszyna], od.plShift as [Zmiana], SUM(od.plAmount) as [Amount] " _
        & "FROM tbOperations o LEFT JOIN tbOperationData od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId " _
        & "LEFT JOIN tbMachine m ON m.machineId=od.plMach " _
        & "WHERE od.plMoment >= @sDate AND od.plMoment <= @eDate AND o.type='r' " _
        & "GROUP BY od.plMoment, z.zfinIndex, m.machineName, od.plShift"
        
Set rs = CreateObject("adodb.recordset")
rs.Open sql, adoConn
Set rs = rs.NextRecordset
Set rs = rs.NextRecordset
If rs.EOF Then
    MsgBox "Brak danych dla wybranego okresu", vbInformation + vbOKOnly, "Brak danych"
Else
    downloadRoastingBatches
    Do Until rs.EOF
        Set nRec = newRecord(rs.Fields("Data"), rs.Fields("Zmiana"))
        If Trim(rs.Fields("Maszyna")) = "RN3000" Then
            nRec.append rs.Fields("Index"), rs.Fields("Amount")
        Else
            nRec.append rs.Fields("Index"), , rs.Fields("Amount")
        End If
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
updateConnection
Exit Sub

err_trap:
MsgBox "Error in ""Update"" of clsRecordsKeepr. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub downloadRoastingBatches()
Dim sql As String
Dim rs As ADODB.Recordset
Dim zfor As clsZfor
Dim bool As Boolean
Dim rn3 As Variant
Dim rn4 As Variant

On Error GoTo err_trap

bool = False

If Not roastingBatches Is Nothing Then
    If roastingBatches.Count = 0 Then
        bool = True
    End If
Else
    Set roastingBatches = New Collection
    bool = True
End If

If bool Then
    connectScada
    sql = "select DISTINCT zlec.MaterialNumber, " _
        & "(SELECT AVG(sub.green) FROM " _
        & "(SELECT TOP(10) z.ILOSC_PALONA as green FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE " _
                                & "JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE " _
                                & "WHERE z.NUMERPIECA=3000 AND z.SUMA_ZIELONEJ>100 AND zlec.MaterialNumber=zl.MaterialNumber ORDER BY z.DTZAPIS DESC) sub) As RN3000, " _
        & "(SELECT AVG(sub.green) FROM " _
        & "(SELECT TOP(10) z.ILOSC_PALONA as green FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE " _
                                & "JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE " _
                                & "WHERE z.NUMERPIECA=4000 AND z.SUMA_ZIELONEJ>100 AND zlec.MaterialNumber=zl.MaterialNumber ORDER BY z.DTZAPIS DESC) sub) As RN4000 " _
    & "FROM ZLECENIA zlec;"
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, scadaConn
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            Set zfor = New clsZfor
            rn3 = rs.Fields("RN3000")
            rn4 = rs.Fields("RN4000")
            If IsNumeric(rn3) Then
                rn3 = Round(rn3, 1)
            End If
            If IsNumeric(rn4) Then
                rn4 = Round(rn4, 1)
            End If
            zfor.initialize rs.Fields("MaterialNumber"), rn3, rn4
            roastingBatches.Add zfor, rs.Fields("MaterialNumber")
            rs.MoveNext
        Loop
    End If
    rs.Close
End If

exit_here:
Set rs = Nothing
disconnectScada
Exit Sub

err_trap:
MsgBox "Error in ""downloadRoastingBatches"" of updater. Error number: " & Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Error"
Resume exit_here

End Sub

Private Function newRecord(aDate As Variant, shift As Integer) As clsRecord
Dim bool As Boolean
Dim cRec As clsRecord
Dim index As String

On Error GoTo err_trap

index = aDate & "_" & shift

For Each cRec In pRecords
    If cRec.Id = index Then
        Set newRecord = cRec
        bool = True
        Exit For
    End If
Next cRec

If Not bool Then
    Set newRecord = New clsRecord
    newRecord.initialize aDate, shift, Me
    pRecords.Add newRecord, index
End If

exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newItem"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Sub display(total As Boolean)
Dim sht As Worksheet
Dim r As clsRecord
Set sht = ThisWorkbook.Sheets("Kontrola upałów")
Dim curDate As Date
Dim r3Total, r4Total, r3Shift, r4Shift As Double

Dim i As Integer
Dim iShift As Integer

r3Total = 0
r4Total = 0

i = 3
sht.Range("A3:H600").Cells.ClearContents

curDate = pDFrom
Do Until curDate > pDTo
    r3Shift = 0
    r4Shift = 0
    If DatePart("h", curDate, vbMonday) = 6 Then
        iShift = 1
    ElseIf DatePart("h", curDate, vbMonday) = 14 Then
        iShift = 2
    Else
        iShift = 3
    End If
    sht.Cells(i, 1) = curDate
    sht.Cells(i, 2) = StrConv(WeekdayName(weekday(curDate, vbMonday), False, vbMonday), vbProperCase)
    sht.Cells(i, 3) = iShift
    If total Then
        r3Shift = getTotalOnDate(curDate, "RN3000")
        r4Shift = getTotalOnDate(curDate, "RN4000")
        sht.Cells(i, 4) = r3Total + r3Shift
        sht.Cells(i, 6) = r4Total + r4Shift
        r3Total = r3Total + r3Shift
        r4Total = r4Total + r4Shift
    Else
        sht.Cells(i, 4) = getTotalOnDate(curDate, "RN3000")
        sht.Cells(i, 6) = getTotalOnDate(curDate, "RN4000")
    End If
    
    i = i + 1
    curDate = DateAdd("h", 8, curDate)
Loop
    
End Sub

Public Function getTotalOnDate(d As Date, roaster As String) As Double
Dim r As clsRecord
Dim amount As Double

For Each r In pRecords
    If DateDiff("s", r.aDate, d) = 0 Then
        If roaster = "RN3000" Then
            amount = amount + r.r3000
        Else
            amount = amount + r.r4000
        End If
    End If
Next r

getTotalOnDate = amount

End Function

Public Sub PrintRecords()
Dim r As clsRecord
Dim lline As String

Debug.Print "Data" & "    " & "Zmiana" & "    " & "RN3000" & "    " & "RN4000"
For Each r In pRecords
    lline = r.aDate & "    " & r.shift & "    " & r.r3000 & "    " & r.r4000
    Debug.Print lline
Next r
End Sub


