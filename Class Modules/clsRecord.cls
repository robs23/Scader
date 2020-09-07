VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private p3000 As Double
Private p4000 As Double
Private pDate As Variant
Private pShift As Integer
Private pRecordsKeeper As clsRecordsKeeper

Public Sub initialize(theDate As Variant, shift As Integer, recordKeeper As clsRecordsKeeper)
pDate = theDate
pShift = shift
Set pRecordsKeeper = recordKeeper
pId = theDate & "_" & shift
End Sub

Public Property Get aDate() As Variant
aDate = pDate
End Property

Public Property Get shift() As Integer
shift = pShift
End Property

Public Property Get r3000() As Double
r3000 = p3000
End Property

Public Property Get r4000() As Double
r4000 = p4000
End Property

Public Sub append(zfor As Long, Optional r3 As Variant, Optional r4 As Variant)
Dim var As Variant

If inCollection(CStr(zfor), pRecordsKeeper.getRoastingBatches) Then
    If Not IsMissing(r3) Then
        var = pRecordsKeeper.getRoastingBatches(CStr(zfor)).rn3000input
        If IsNull(var) Then var = 340
        p3000 = p3000 + Round(r3 / var, 0)
    End If
    If Not IsMissing(r4) Then
        var = pRecordsKeeper.getRoastingBatches(CStr(zfor)).rn4000input
        If IsNull(var) Then var = 425
        p4000 = p4000 + Round(r4 / var, 0)
    End If
End If
End Sub

Public Property Get Id() As String
Id = pId
End Property

