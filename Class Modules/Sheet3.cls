VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private dbl As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
Dim val As Variant
Dim con As Variant
Dim sht As Worksheet
Dim c As Range
Dim pal As Double

Set sht = ThisWorkbook.Sheets("Kontrola upałów")

If Not syncRunning Then
    If dbl Then
        dbl = False
    Else
        If Target.Cells.Count = 1 Then
            If Len(Target) > 0 Then
                If Not Intersect(Target, sht.Range("A:A")) Is Nothing Then
                    Set c = Target
                    If Not IsDate(c) Then
                        dbl = True
                        MsgBox "Podana wartość nie jest datą. Podaj wartość w formacie daty np. """ & Date & """", vbExclamation + vbOKOnly, "Błąd daty"
                        c = ""
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

