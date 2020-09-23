VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICompareClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ICompareClient

' Public interface for the cStableSorter class to
' be Implemented by the client.

Public Sub Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eiCompare, ByVal Percent As Long, Cancel As Boolean)
'
End Sub

' In the client:

'Private cSorter As cStableSorter
'Implements ICompareClient

'Set cSorter = New cStableSorter
'cSorter.Attach Me

'ReDim idxA(lb To ub) As Long

'cSorter.Sort idxA, lb, ub

'Private Sub ICompareClient_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eiCompare, ByVal Percent As Long, Cancel As Boolean)
'    pgbProgress.Value = Percent
'    If srcA(ThisIdx) > srcA(ThanIdx) Then
'        Result = Greater
'    ElseIf srcA(ThisIdx) < srcA(ThanIdx) Then
'        Result = Lesser
'    Else
'        Result = Equal
'    End If
'    DoEvents
'    If mCancelFlag Then Cancel = True
'End Sub

'For i = lb To ub
'    item = srcA(idxA(i))
'Next

'cSorter.Detach
'Set cSorter = Nothing
