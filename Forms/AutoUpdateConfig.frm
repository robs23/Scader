VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoUpdateConfig 
   Caption         =   "Konfiguracja auto odświeżania"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "AutoUpdateConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoUpdateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnStart_Click()
Dim cycle As Integer

If validate Then
    refreshCycle = Me.txtCycle.Value
    refreshRange = Me.txtRange.Value
    If isOnTimeSet Then
        StopAutoUpdate
    End If
    RefreshCharts
End If


End Sub


Private Function validate() As Boolean
Dim pass As Boolean

pass = False

If Not IsNumeric(Me.txtCycle.Value) Or Not IsNumeric(Me.txtRange.Value) Then
    MsgBox "Zarówno cykl jak i zakres powinny być wartościami liczbowymi większymi od 0! Proszę poprawić", vbCritical + vbOKOnly, "Błąd"
Else
    If Me.txtCycle.Value <= 0 Or Me.txtRange.Value <= 0 Then
        MsgBox "Zarówno cykl jak i zakres powinny być wartościami liczbowymi większymi od 0! Proszę poprawić", vbCritical + vbOKOnly, "Błąd"
    Else
        pass = True
    End If
End If

validate = pass

End Function
