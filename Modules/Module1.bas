Attribute VB_Name = "Module1"
Public conn As ADODB.Connection 'connection to scada
Public blends As New Collection
Public isError As Boolean
Public syncRunning As Boolean
Public records As New Collection
Public refreshCycle As Integer
Public refreshRange As Integer
Public nextUpdateTime As Date
Public isOnTimeSet As Boolean


Public Sub Connection(StartDate As Date, endDate As Date, Optional roaster As Variant, Optional blends As Variant, Optional exclude As Variant, Optional divOnRoasters As Variant)

    On Error GoTo err_trap
    
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim rcrds As ADODB.Recordset
    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Dim blendString As String
    Dim excludeString As String
    Dim i As Integer

    conn.Provider = "SQLOLEDB"
    conn.connectionString = scadaConnectionString
    conn.Open
    conn.CommandTimeout = 90
    
    If Not IsMissing(blends) Then
        If Not isArrayEmpty(blends) Then
            blendString = " AND ("
            For i = LBound(blends) To UBound(blends)
                If i = LBound(blends) Then
                    blendString = blendString & "zl.MaterialNumber = " & blends(i)
                Else
                    blendString = blendString & " OR zl.MaterialNumber = " & blends(i)
                End If
            Next i
            blendString = blendString & ")"
        End If
    End If
    If Not IsMissing(exclude) Then
        If Not isArrayEmpty(exclude) Then
            excludeString = " AND ("
            For i = LBound(exclude) To UBound(exclude)
                If i = LBound(exclude) Then
                    excludeString = excludeString & "zl.MaterialNumber <> " & exclude(i)
                Else
                    excludeString = excludeString & " AND zl.MaterialNumber <> " & exclude(i)
                End If
            Next i
            excludeString = excludeString & ")"
        End If
    End If

    SQLstr = "select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT" _
    & " from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)" _
    & " Where (z.DTZAPIS Between ('" & Format(StartDate, "yyyy-mm-dd hh:mm") & "') AND ('" & Format(endDate, "yyyy-mm-dd hh:mm") & "'))"
    If Not IsMissing(roaster) Then SQLstr = SQLstr & " AND z.NUMERPIECA = " & roaster
    If blendString <> "" Then SQLstr = SQLstr & blendString
    If excludeString <> "" Then SQLstr = SQLstr & excludeString
    SQLstr = SQLstr & " ORDER BY z.DTZAPIS;"

'
   'wykonanie zapytania i przypisanie wyniku do zmiennej rekordow
   Set rcrds = conn.Execute(SQLstr)
 
   i = 1
   
   With ThisWorkbook.Sheets("Arkusz1")
        
        'zapisywanie wyniku zapytania w arkuszu - iteracja zestawu rekordow
        .Cells(i, 1) = "Piec"
        .Cells(i, 2) = "Kawa zielona"
        .Cells(i, 3) = "Uprażono"
        .Cells(i, 4) = "Data"
        .Cells(i, 5) = "Zlecenie"
        .Cells(i, 6) = "ZFOR"
        .Cells(i, 7) = "Nazwa"
        .Cells(i, 8) = "Ubytek [%]"
        i = 2
        Do While Not rcrds.EOF
            .Cells(i, 1) = rcrds("NUMERPIECA")
            .Cells(i, 2) = rcrds("SUMA_ZIELONEJ")
            .Cells(i, 3) = rcrds("ILOSC_PALONA")
            .Cells(i, 4) = rcrds("DTZAPIS")
            .Cells(i, 4).NumberFormat = "dd-mm-yyyy hh:mm:ss"
            .Cells(i, 5) = rcrds("OrderNumber")
            .Cells(i, 6) = rcrds("MaterialNumber")
            .Cells(i, 7) = rcrds("NAZWARECEPT")
             If Not IsNull(rcrds("ILOSC_PALONA")) And Not IsNull(rcrds("SUMA_ZIELONEJ")) And Not rcrds("SUMA_ZIELONEJ") = 0 Then
                 .Cells(i, 8) = 1 - (rcrds("ILOSC_PALONA") / rcrds("SUMA_ZIELONEJ"))
                 .Cells(i, 8).NumberFormat = "0.00%"
             End If
            rcrds.MoveNext
            i = i + 1
        Loop
   End With
   
 
   'zakonczenie polaczenia
   rcrds.Close
   Dim div As Boolean
   div = False
   If Not IsMissing(divOnRoasters) Then
        div = divOnRoasters
   End If
    If UserForm1.cboxTimeDiff.Value = True Or div Then
        divideOnRoasters StartDate, endDate, "3000"
        divideOnRoasters StartDate, endDate, "4000"
    End If
   
   conn.Close
   Set rcrds = Nothing
    Set conn = Nothing
If UserForm1.cboxGraph.Value = True Then
    createGraph
End If

exit_here:
Exit Sub

err_trap:
isOnTimeSet = False
MsgBox "Błąd w Connection: " & Err.Description, vbCritical + vbOKOnly, "Błąd"
Resume exit_here

End Sub

Sub startForm()
UserForm1.Show
End Sub

Public Sub adjustCharts()
Dim chrt As ChartObject
Dim sht As Worksheet
Dim i As Integer
Dim n As Integer
Dim lastCell As Range
Dim aStr As String
Dim v() As String
Dim x() As String
Dim xStr As String
Dim rXVal As Range
Dim counter As Integer
Dim rangeDef As String

On Error GoTo err_trap

For i = 0 To 1
    If i = 0 Then
        Set sht = ThisWorkbook.Sheets("RN3000")
        Set chrt = ThisWorkbook.Sheets("Wykresy").ChartObjects(1)
    Else
        Set sht = ThisWorkbook.Sheets("RN4000")
        Set chrt = ThisWorkbook.Sheets("Wykresy").ChartObjects(2)
    End If
    Set lastCell = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious)
    If Not lastCell Is Nothing Then
        If lastCell.row > 1 Then
            'we can adjust the chart
            For n = lastCell.row To 1
                'check if there are rows with 0% of lost coffee and delete themm
                If sht.Range("K" & n) = 0 Then
                    sht.Rows(n).Delete
                End If
            Next n
            Set lastCell = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious)
            If Not lastCell Is Nothing Then
                If lastCell.row > 1 Then
                    'adjust charts
                    For Each srs In chrt.Chart.SeriesCollection
                        v = Split(srs.Formula, "$", , vbTextCompare)
                        If UBound(v) >= 4 Then
                            x = Split(v(4), ",", , vbTextCompare)
                            xStr = "$" & x(0) & ","
                            srs.Formula = Replace(srs.Formula, xStr, "$" & lastCell.row & ",", , , vbTextCompare)
'                            Set rXVal = Range(Split(srs.Formula, ",")(1))
'                            srs.XValues = rXVal
again:
                        End If
                    Next srs
                End If
            End If
        End If
    End If
Next i

exit_here:
ThisWorkbook.Sheets("Wykresy").Range("D8") = ""
If ActiveWorkbook.Name = ThisWorkbook.Name Then
    ThisWorkbook.Sheets("Wykresy").Activate
End If
Exit Sub

err_trap:
isOnTimeSet = False
counter = counter + 1
If counter < 10 Then
    GoTo again
Else
    Resume exit_here
End If

End Sub

Public Function rangeDef(rngStr As Variant) As Range
Dim v() As String
Dim shtName As String
Dim rng As Range

v = Split(rngStr, "!", , vbTextCompare)

shtName = Left(v(0), Len(v(0)) - 1)
shtName = Right(shtName, Len(shtName) - 1)

Set rng = ThisWorkbook.Sheets(shtName).Range(v(1))

Set rangeDef = rng

End Function

Public Function isArrayEmpty(parArray As Variant, Optional dimension As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
    If IsMissing(dimension) Then
        If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    Else
        If UBound(parArray, dimension) < LBound(parArray, dimension) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    End If
End Function

Public Sub createGraph(Optional timeOnX As Variant)
Dim i As Integer
Dim lineChart As ChartObject
Dim chD() As Variant 'chartData
Dim found As Boolean
Dim n As Integer
Dim q As Integer
Dim blend3() As Long 'blends on RN3000
Dim blend4() As Long 'blends on RN4000
Dim b As Long 'single blend
Dim r As Long 'roaster
Dim rn4000max As Double
Dim rn4000min As Double
Dim rn3000max As Double
Dim rn3000min As Double
Dim x() As Long
Dim y() As Double
Dim row As Long


For i = 2 To 10000
    If ThisWorkbook.Sheets(1).Cells(i, 6) > 0 Then
        r = ThisWorkbook.Sheets(1).Cells(i, 1)
        b = ThisWorkbook.Sheets(1).Cells(i, 6)
        If r = 3000 Then
            If isArrayEmpty(blend3) Then
                ReDim blend3(0) As Long
                blend3(0) = b 'blend
            Else
                found = False
                For n = LBound(blend3) To UBound(blend3)
                    If blend3(n) = b Then
                        found = True
                        Exit For
                    End If
                Next n
                If found = False Then
                    ReDim Preserve blend3(UBound(blend3) + 1) As Long
                    blend3(UBound(blend3)) = b
                End If
            End If
        ElseIf r = 4000 Then
            If isArrayEmpty(blend4) Then
                ReDim blend4(0) As Long
                blend4(0) = b 'blend
            Else
                found = False
                For n = LBound(blend4) To UBound(blend4)
                    If blend4(n) = b Then
                        found = True
                        Exit For
                    End If
                Next n
                If found = False Then
                    ReDim Preserve blend4(UBound(blend4) + 1) As Long
                    blend4(UBound(blend4)) = b
                End If
            End If
        End If
    Else
        Exit For
    End If
Next i

If Not isArrayEmpty(blend3) Then
    rn3000max = 0
    rn3000min = 50
    Set rng = ThisWorkbook.Sheets(1).Range("I10:Y30")
    Set lineChart = ThisWorkbook.Sheets(1).ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, Top:=rng.Top, Height:=rng.Height)
    With lineChart
        .Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, Title:="RN3000"
        .Name = "RN3000"
        For i = LBound(blend3) To UBound(blend3)
            For row = 2 To 10000
                If ThisWorkbook.Sheets(1).Cells(row, 1) = 3000 And ThisWorkbook.Sheets(1).Cells(row, 6) = blend3(i) Then
                    If isArrayEmpty(y) Then
                        ReDim y(0) As Double
                        y(0) = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 > rn3000max Then rn3000max = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(1).Cells(row, 8) * 100 < rn3000min Then rn3000min = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                    Else
                        ReDim Preserve y(UBound(y) + 1) As Double
                        y(UBound(y)) = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 > rn3000max Then rn3000max = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(1).Cells(row, 8) * 100 < rn3000min Then rn3000min = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                    End If
                ElseIf ThisWorkbook.Sheets(1).Cells(row, 1) = 0 Then
                    With .Chart
                     .SeriesCollection.NewSeries
                        With .SeriesCollection(i + 1)
                                .Name = blend3(i) & " " & getBlendName(blend3(i))
                                .values = y
        '                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
        '                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                            .MarkerStyle = xlMarkerStyleNone
                            Erase y
        '                    .ApplyDataLabels
        ''                    .DataLabels.Select
                        End With
                        If rn3000min < 10 Then
                            rn3000min = 10
                        Else
                            rn3000min = rn3000min - 1
                        End If
                        If rn3000max > 20 Then
                            rn3000max = 20
                        Else
                            rn3000max = rn3000max + 1
                        End If
                        .Axes(xlValue).MinimumScale = Int(rn3000min)
                        .Axes(xlValue).MaximumScale = Int(rn3000max)
                    End With
                    Exit For
                End If
            Next row
        Next i
    End With
    Set lineChart = Nothing
End If
If Not isArrayEmpty(blend4) Then
    Set rng = ThisWorkbook.Sheets(1).Range("I35:Y55")
    Set lineChart = ThisWorkbook.Sheets(1).ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, Top:=rng.Top, Height:=rng.Height)
    With lineChart
        .Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, Title:="RN4000"
        .Name = "RN4000"
        For i = LBound(blend4) To UBound(blend4)
            For row = 2 To 10000
                If ThisWorkbook.Sheets(1).Cells(row, 1) = 4000 And ThisWorkbook.Sheets(1).Cells(row, 6) = blend4(i) Then
                    If isArrayEmpty(y) Then
                        ReDim y(0) As Double
                        y(0) = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 > rn4000max Then rn4000max = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(1).Cells(row, 8) * 100 < rn4000min Then rn4000min = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                    Else
                        ReDim Preserve y(UBound(y) + 1) As Double
                        y(UBound(y)) = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 > rn4000max Then rn4000max = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(1).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(1).Cells(row, 8) * 100 < rn4000min Then rn4000min = ThisWorkbook.Sheets(1).Cells(row, 8) * 100
                    End If
                ElseIf ThisWorkbook.Sheets(1).Cells(row, 1) = 0 Then
                    With .Chart
                     .SeriesCollection.NewSeries
                        With .SeriesCollection(i + 1)
                                .Name = blend4(i) & " " & getBlendName(blend4(i))
                                .values = y
        '                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
        '                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                            .MarkerStyle = xlMarkerStyleNone
                            Erase y
        '                    .ApplyDataLabels
        ''                    .DataLabels.Select
                        End With
                        If rn4000min < 10 Then
                            rn4000min = 10
                        Else
                            rn4000min = rn4000min - 1
                        End If
                        If rn4000max > 20 Then
                            rn4000max = 20
                        Else
                            rn4000max = rn4000max + 1
                        End If
                        .Axes(xlValue).MinimumScale = Int(rn4000min)
                        .Axes(xlValue).MaximumScale = Int(rn4000max)
                    End With
                    Exit For
                End If
            Next row
        Next i
    End With
    Set lineChart = Nothing
End If
'If ThisWorkbook.Sheets(1).ChartObjects.Count > 0 Then
'    ThisWorkbook.Sheets(1).ChartObjects.Delete
'End If
'
'n = 0
'
'For i = 2 To 10000
'    If ThisWorkbook.Sheets(1).Cells(i, 6) = blend Then
'    If isArrayEmpty(chD) Then
'        ReDim chD(0, 3) As Variant
'        chD(0, 0) = n
'    Else
'        ReDim chD(UBound(chD, 1) + 1, 3) As Variant
'    End If
'Next i
'
'

'


End Sub

Public Sub applyCustomPointLabels(seriesName As String, ch As ChartObject, Optional values As Variant)
Dim srs As Series, rng As Range, lbl As DataLabel
Dim iLbl As Long, nLbls As Long
Dim pnt As Point

Set srs = ch.Chart.SeriesCollection(seriesName)

If Not srs Is Nothing Then
    For Each pnt In srs.Points
        pnt.HasDataLabel = True
        Set lbl = pnt.DataLabel
        With lbl
            .Text = "Dupa"
            .Position = xlLabelPositionRight
        End With
        Set lbl = Nothing
    Next pnt
End If
Set srs = Nothing
End Sub

Sub xx()
applyCustomPointLabels "30001565", ThisWorkbook.Sheets(1).ChartObjects("RN4000")
End Sub

Public Sub PROD_PODSUMUJ() 'rng As Range, id As Integer, value As Integer)
Dim i As Integer
Dim idCell As Range
Dim valCell As Range
Dim val As Double
Dim index As Variant
Dim oW As Worksheet
Dim Target As Range
Dim n As Integer

Target = Application.ActiveCell
Set oW = rng.Worksheet
rng.Sort key1:=oW.Cells(rng.Column + Id - 1), order1:=xlAscending
For i = rng.row To rng.Height + rng.row
    idCell = oW.Cells(i, rng.Column + Id - 1)
    If Not IsEmpty(idCell) Then
        valCell = oW.Cells(i, rng.Column + Value - 1)
        If IsNumeric(valCell) Then
            If idCell.Value = index Then
                val = val + idCell.Value
            Else
                Target.Worksheet.Cells(Target.row + n, Target.Column) = index
                Target.Worksheet.Cells(Target.row + n, Target.Column + 1) = val
                index = idCell.Value
                val = 0
                n = n + 1
            End If
        End If
        
    End If
Next i

End Sub

Public Function getBlendName(blend As Long) As String
Dim i As Integer
For i = 2 To 10000
    If ThisWorkbook.Sheets(1).Cells(i, 6) = blend Then
        getBlendName = ThisWorkbook.Sheets(1).Cells(i, 7)
        Exit For
    ElseIf ThisWorkbook.Sheets(1).Cells(i, 6) = "" Then
        Exit For
    End If
Next i
End Function


Public Sub divideOnRoasters(StartDate As Date, endDate As Date, r As String)
Dim SQLstr As String
Dim rs As ADODB.Recordset
Dim lim As Variant
Dim prevT As Date

On Error GoTo err_trap

isError = False

downloadBlends
If Not isError Then
    SQLstr = "select DISTINCT z.ID_ZLECENIA_PALONA, z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT" _
        & " from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)" _
        & " Where (z.DTZAPIS Between ('" & Format(StartDate, "yyyy-mm-dd hh:mm") & "') AND ('" & Format(endDate, "yyyy-mm-dd hh:mm") & "')) AND z.NUMERPIECA = " & r
        SQLstr = SQLstr & " ORDER BY z.DTZAPIS;"
    
    '
    'wykonanie zapytania i przypisanie wyniku do zmiennej rekordow
    Set rs = CreateObject("adodb.recordset")
    rs.Open SQLstr, conn
    If Not rs.EOF Then
        With ThisWorkbook.Sheets("RN" & r)
            .Range("A1") = "ID"
            .Range("B1") = "Piec"
            .Range("C1") = "Kawa zielona"
            .Range("D1") = "Uprażono"
            .Range("E1") = "Data"
            .Range("F1") = "Czas postoju"
            .Range("G1") = "Limit"
            .Range("H1") = "Zlecenie"
            .Range("I1") = "ZFOR"
            .Range("J1") = "Nazwa"
            .Range("K1") = "Ubytek [%]"
            .Range("L1") = "Komentarz"
            rs.MoveFirst
            i = 1
            Do Until rs.EOF
                i = i + 1
                .Range("A" & i) = rs.Fields("ID_ZLECENIA_PALONA")
                .Range("B" & i) = rs.Fields("NUMERPIECA")
                .Range("C" & i) = rs.Fields("SUMA_ZIELONEJ")
                .Range("D" & i) = rs.Fields("ILOSC_PALONA")
                .Range("E" & i) = rs.Fields("DTZAPIS")
                .Range("E" & i).NumberFormat = "dd-mm-yyyy hh:mm:ss"
                If Abs(DateDiff("yyyy", Now, prevT)) < 20 Then
                    .Range("F" & i) = CDate(rs.Fields("DTZAPIS") - prevT)
                    .Range("F" & i).NumberFormat = "hh:mm:ss"
                    lim = blendBreak(rs.Fields("MaterialNumber"))
                    If DateDiff("s", prevT, rs.Fields("DTZAPIS")) >= lim Then
                        'if break shorter than 25 minutes, color in yellow
                        If DateDiff("s", prevT, rs.Fields("DTZAPIS")) >= 1500 Then
                            .Range("F" & i).Interior.Color = RGB(255, 153, 153)
                        Else
                           .Range("F" & i).Interior.Color = RGB(255, 255, 102)
                        End If
                    End If
                    .Range("G" & i) = sek2Date(CLng(lim))
                    .Range("G" & i).NumberFormat = "hh:mm:ss"
                End If
                .Range("H" & i) = rs.Fields("OrderNumber")
                .Range("I" & i) = rs.Fields("MaterialNumber")
                .Range("J" & i) = rs.Fields("NAZWARECEPT")
                 If Not IsNull(rs.Fields("ILOSC_PALONA")) And Not IsNull(rs.Fields("SUMA_ZIELONEJ")) And Not rs.Fields("SUMA_ZIELONEJ") = 0 Then
                    .Range("K" & i) = 1 - (rs.Fields("ILOSC_PALONA") / rs.Fields("SUMA_ZIELONEJ"))
                    .Range("K" & i).NumberFormat = "0.00%"
                End If
                prevT = rs.Fields("DTZAPIS")
                rs.MoveNext
            Loop
            .Range("A:J").AutoFilter
       End With
    End If
End If

exit_here:
Exit Sub

err_trap:
MsgBox "Błąd w divideOnRoasters: " & Err.Description, vbOKOnly + vbCritical, "Błąd"
Resume exit_here

End Sub

Public Sub stackResults()
Dim i As Integer
Dim n As Integer
Dim c As Range
Dim rng As Range
Dim Id As Long
Dim lastRow As Long
Dim kom As String
Dim sht As Worksheet
Dim dest As Worksheet
Dim theRow As Long
Dim lastRow2 As Long
Dim res As VbMsgBoxResult

On Error GoTo err_trap


For i = 3000 To 4000 Step 1000
    Set sht = ThisWorkbook.Sheets("RN" & i)
    Set dest = ThisWorkbook.Sheets("Historia")
    lastRow2 = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    
    For n = 2 To lastRow2
        kom = sht.Range("L" & n)
        theRow = 0
        If Len(kom) > 0 Then
            'new coment has been found
            Id = sht.Range("A" & n)
            If Not Id = 0 Then
'                lastRow = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                Set rng = dest.Range("A:A")
                theRow = rng.Find(Id, searchorder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).row
                If theRow > 0 Then
                    If kom <> dest.Range("K" & theRow) Then
                        'different comments for the same roasting batch
                        res = MsgBox("Dla upału " & Id & " istnieje już komentarz i jest on inny niż ten, który właśnie próbujesz zapisać. Istniejący komentarz: """ & dest.Range("K" & theRow) & """, nowy komentarz: """ & kom & """. Czy zastąpić oryginalny komentarz tym nowym?", vbQuestion + vbYesNo, "Komentarz już istnieje")
                        If res = vbYes Then
                            'replace the comment
                            sht.Range("A" & n & ":F" & n).Copy dest.Range("A" & theRow & ":F" & theRow)
                            sht.Range("H" & n & ":L" & n).Copy dest.Range("G" & theRow & ":K" & theRow)
                        End If
                    End If
                End If
            End If
        End If
    Next n
Next i
MsgBox "Zmiany zostały zapisane w historii", vbOKOnly + vbInformation, "Zapisano"

exit_here:
Set sht = Nothing
Set dest = Nothing
Set rng = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    'completely new comment
    lastRow = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    sht.Range("A" & n & ":F" & n).Copy dest.Range("A" & lastRow + 1 & ":F" & lastRow + 1)
    sht.Range("H" & n & ":L" & n).Copy dest.Range("G" & lastRow + 1 & ":K" & lastRow + 1)
    Resume Next
Else
    MsgBox "Error in stackResults. Error nubmer:  " & Err.Number & ", " & Err.Description
    Resume exit_here
End If
End Sub

Public Sub downloadBlends()
Dim i As Integer
Dim n As Integer
Dim ind As Long
Dim t As Date
Dim b As clsBlend

On Error GoTo err_trap

n = blends.Count
Do While blends.Count > 0
    blends.Remove n
    n = n - 1
Loop

With ThisWorkbook.Sheets("Limity")
    For i = 2 To 1000
        ind = .Range("A" & i)
        If ind = 0 Then
            Exit For
        Else
            Set b = New clsBlend
            b.blendIndex = ind
            b.blendName = .Range("B" & i)
            b.sekBreak = min2Sek(CStr(.Range("C" & i)))
            blends.Add b, CStr(ind)
        End If
    Next i
End With

exit_here:
Exit Sub

err_trap:
isError = True
If Err.Number = 457 Then
    MsgBox "W arkuszu ""Limity"" numer " & ind & " występuje wielokrotnie. Usuń wszystkie powtórzenia tego indeksu", vbOKOnly + vbCritical, "Dubel"
Else
    MsgBox "Error in downloadBlends. Error number: " & Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Błąd"
End If
Resume exit_here

End Sub

Public Function min2Sek(minStr As Date) As Long
Dim v() As String
v() = Split(minStr, ":", , vbTextCompare)
If UBound(v) > 0 Then
    min2Sek = CLng(v(0)) * 3600 + (CLng(v(1)) * 60) + CLng(v(2))
End If

End Function

Private Function blendBreak(blendNumber As Long) As Long
    Dim val As Long

On Error GoTo err_trap

val = 720

val = blends(CStr(blendNumber)).sekBreak

exit_here:
blendBreak = val
Exit Function

err_trap:
Resume exit_here

End Function

Public Function sek2Date(sek As Long) As Date
Dim m As Long
Dim h As Long
Dim s As Long

If Not sek = 0 Then
    h = Int(sek / 3600)
    m = Int((sek - (h * 3600)) / 60)
    s = Int(sek - (h * 3600) - (m * 60))
    sek2Date = TimeSerial(h, m, s)
End If
End Function



'Public Sub updateRoastingStats()
'Dim lastRow As Long
'Dim sht As Worksheet
'Dim iShift As Integer
'Dim i As Integer
'Dim counter As Integer
'Dim curDate As Date
'Dim lastDate As Date
'Dim sql As String
'Dim r3Total As Double
'Dim r4Total As Double
'Dim rs As ADODB.Recordset
'Dim nRec As clsRecord
'Dim sDate As Date
'
'On Error GoTo err_trap
'
'Application.ScreenUpdating = False
'Application.StatusBar = "Wczytuję dane.. Proszę czekać.. "
'Application.Cursor = xlWait
'syncRunning = True
''getRoastingBatches
'
'For Each nRec In records
'    records.Remove nRec.Id
'Next nRec
'
'Set sht = ThisWorkbook.Sheets("Kontrola upałów")
'
'lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
'
'If lastRow > 2 Then
'    updateConnection
'    i = lastRow
'    counter = 0
'    Do Until i <= 2
'        'if there are less than 10 records, finish when reached top of the sheet
'        If counter >= 10 Then
'            'finish after 10 last records
'            Exit Do
'        Else
'            If sht.Range("E" & i) <> "" Or sht.Range("G" & i) <> "" Then
'                'finish when first filled record is found
'                Exit Do
'            Else
'                curDate = sht.Range("A" & i)
'                sDate = curDate
'            End If
'        End If
'        counter = counter + 1
'        i = i - 1
'    Loop
'    'ok, we've got date to strart with
'
'    sql = "SELECT TOP(1) od.plMoment as lastDate " _
'        & "FROM tbOperations o LEFT JOIN tbOperationData od ON od.operationId=o.operationId " _
'        & "WHERE o.type = 'r' " _
'        & "ORDER BY od.plMoment DESC"
'
'    Set rs = CreateObject("adodb.recordset")
'    rs.Open sql, adoConn
'    If rs.EOF Then
'        rs.Close
'        GoTo err_trap
'    Else
'        rs.MoveFirst
'        lastDate = rs.Fields("lastDate")
'        rs.Close
'    End If
'    'curDate = DateAdd("h", 14, curDate)
'    Do Until curDate >= lastDate
'        r3Total = 0
'        r4Total = 0
'        sql = "DECLARE @sDate as datetime; " _
'            & "DECLARE @eDate as datetime; " _
'            & "SELECT @sDate = '" & curDate & "'; " _
'            & "SELECT @eDate = '" & DateAdd("d", 1, curDate) & "'; " _
'            & "SELECT z.zfinIndex,z.zfinName, CONVERT(date,od.plMoment,103) as plDate,od.plShift, m.machineName as Maszyna, SUM(od.plAmount) as KG " _
'            & "FROM tbOperations o LEFT JOIN tbOperationDataHistory od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId " _
'            & "WHERE od.plMoment >= @sDate AND od.plMoment < @eDate AND o.type = 'r' AND od.operDataVer = " _
'            & "(SELECT TOP (1) sub.operDataVer FROM(SELECT oh.operDataVer, ov.createdOn, CAST(COUNT(oh.operationId) as float) / CAST((SELECT COUNT(ohAll.operationId) FROM tbOperationData ohAll WHERE ohAll.plMoment BETWEEN @sDate AND @eDate) AS float) as perc " _
'            & "FROM tbOperationDataHistory oh JOIN tbOperationDataVersions ov ON ov.operDataVerId=oh.operDataVer " _
'            & "WHERE oh.plMoment BETWEEN @sDate AND @eDate " _
'            & "GROUP BY oh.operDataVer, ov.createdOn " _
'            & "HAVING CAST(COUNT(oh.operationId) as float) / CAST((SELECT COUNT(ohAll.operationId) FROM tbOperationData ohAll WHERE ohAll.plMoment BETWEEN @sDate AND @eDate) AS float) > 0.8) sub " _
'            & "WHERE sub.createdOn <= DATEADD(HOUR,10,@sDate) " _
'            & "ORDER BY sub.createdOn DESC) " _
'            & "GROUP BY z.zfinIndex, z.zfinName, CONVERT(date,od.plMoment,103),od.plShift, m.machineName " _
'            & "ORDER BY m.machineName, od.plShift"
'
'        rs.Open sql, adoConn
'        If Not rs.EOF Then
'            rs.MoveFirst
'            Do Until rs.EOF
'                'Debug.Print rs.Fields("zfinIndex")
'                Set nRec = newRecord(rs.Fields("plDate"), rs.Fields("plShift"))
'                If Trim(rs.Fields("Maszyna")) = "RN3000" Then
'                    nRec.append rs.Fields("zfinIndex"), rs.Fields("KG")
'                Else
'                    nRec.append rs.Fields("zfinIndex"), , rs.Fields("KG")
'                End If
'                rs.MoveNext
'            Loop
'        End If
'        rs.Close
'
'        curDate = DateAdd("d", 1, curDate)
'    Loop
'    For Each nRec In records
'        If i <= 3 Then i = 3
'        If (nRec.aDate = sDate And nRec.shift = 3) Or (nRec.aDate > sDate) Then
'            sht.Range("A" & i) = CDate(nRec.aDate)
'            sht.Range("B" & i) = StrConv(WeekdayName(weekday(CDate(nRec.aDate), vbMonday), False, vbMonday), vbProperCase)
'            sht.Range("C" & i) = nRec.shift
'            sht.Range("D" & i) = nRec.r3000
'            sht.Range("F" & i) = nRec.r4000
'            i = i + 1
'        End If
'    Next
'    If weekday(lastDate, vbMonday) < 7 Then
'        curDate = DateAdd("h", 8, lastDate)
'        Do Until weekday(curDate, vbMonday) = 7 And DatePart("h", curDate, vbMonday) = 22
'            If DatePart("h", curDate, vbMonday) = 6 Then
'                iShift = 1
'            ElseIf DatePart("h", curDate, vbMonday) = 14 Then
'                iShift = 2
'            Else
'                iShift = 3
'            End If
'            sht.Range("A" & i) = curDate
'            sht.Range("B" & i) = StrConv(WeekdayName(weekday(curDate, vbMonday), False, vbMonday), vbProperCase)
'            sht.Range("C" & i) = iShift
'            sht.Range("D" & i) = 0
'            sht.Range("F" & i) = 0
'            i = i + 1
'            curDate = DateAdd("h", 8, curDate)
'        Loop
'    End If
'End If
'
'exit_here:
'Application.ScreenUpdating = True
'Application.StatusBar = ""
'Application.Cursor = xlDefault
'syncRunning = False
'closeConnection
'Exit Sub
'
'err_trap:
'MsgBox "Error in ""updateRoastingStats"". Error number: " & Err.Number & ", " & Err.Description
'Resume exit_here
'
'End Sub

Public Sub updateGrinding()
Dim sql As String

getWeek.Show

End Sub

Public Function inCollection(ind As String, col As Collection) As Boolean
Dim v As Variant
Dim isError As Boolean

isError = False

On Error GoTo err_trap

Set v = col(ind)

exit_here:
If isError Then
    inCollection = False
Else
    inCollection = True
End If
Exit Function

err_trap:
isError = True
Resume exit_here


End Function

Private Function newRecord(aDate As Variant, shift As Integer, recKeeper As clsRecordsKeeper) As clsRecord
Dim bool As Boolean
Dim cRec As clsRecord
Dim index As String

On Error GoTo err_trap

index = aDate & "_" & shift

For Each cRec In recKeeper.getRoastingBatches
    If cRec.Id = index Then
        Set newRecord = cRec
        bool = True
        Exit For
    End If
Next cRec

If Not bool Then
    Set newRecord = New clsRecord
    newRecord.initialize aDate, shift, recKeeper
    recKeeper.getRoastingBatches.Add newRecord, index
End If

exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newItem"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Sub configureAutoRefresh()
AutoUpdateConfig.Show
End Sub


Public Sub startRoastingForm()
getWeekR.Show

End Sub

Public Sub RefreshCharts()
Dim chrt As ChartObject
Dim Tsht As Worksheet
Dim sht3 As Worksheet
Dim sht4 As Worksheet
Dim isError As Boolean

On Error GoTo err_trap

AutoUpdateConfig.Hide
Application.StatusBar = "Odświeżam wykres.."
Application.ScreenUpdating = False
ThisWorkbook.Sheets("Arkusz1").Cells.Clear
ThisWorkbook.Sheets("RN3000").Cells.ClearContents
ThisWorkbook.Sheets("RN4000").Cells.ClearContents
Connection DateAdd("h", refreshRange * -1, Now), Now, , , , True
adjustCharts

Set Tsht = ThisWorkbook.Sheets("Wykresy")
Set sht3 = ThisWorkbook.Sheets("RN3000")
Set sht4 = ThisWorkbook.Sheets("RN4000")

'remove charts in Wykesy tab
'For Each chrt In Tsht.ChartObjects
'    chrt.Delete
'Next chrt
'sht3.ChartObjects(1).Select
'sht3.ChartObjects(1).Copy
'Tsht.Paste Tsht.Range("B1:Q18")
'
'sht4.ChartObjects(1).Select
'sht4.ChartObjects(1).Copy
'Tsht.Paste Tsht.Range("B19:Q36")

'configure timer for another go
nextUpdateTime = DateAdd("n", refreshCycle, Now)
Application.OnTime nextUpdateTime, "RefreshCharts"

exit_sub:
Application.ScreenUpdating = True
Application.StatusBar = "Odświeżono: " & Now
isOnTimeSet = True
Exit Sub

err_trap:
isOnTimeSet = False
isError = True
MsgBox "Pojawił się błąd przy próbie odświeżenia wykresu: " & Err.Description, vbOKOnly + vbCritical, "Błąd"
Resume exit_sub

End Sub

Public Sub StopAutoUpdate()
If isOnTimeSet Then
    Application.OnTime nextUpdateTime, "RefreshCharts", , False
End If
End Sub

