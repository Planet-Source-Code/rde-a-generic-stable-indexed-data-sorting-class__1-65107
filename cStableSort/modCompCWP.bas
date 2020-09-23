Attribute VB_Name = "modCompCWP"
Option Explicit

Private Declare Function GetInput Lib "user32" Alias "GetInputState" () As Long
'&&&&&&&&&&&&&& Demo code only &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private eCase As VbCompareMethod
Private bCancelFlag As Boolean
Private sA() As String                'attempt to make level performance test
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

' CWP_Compare

' This callback function is required by the cStableSortCWP class.

' The CallWindowProc callback provides greater speed over Event raising.

' You must modify the code below within the CWP_Compare callback, and must
' assign to the functions return value an eCompareResult enumeration value.

' Your callback function MUST conform to the following declaration signiture
' in a standard module:

Function CWP_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, ByVal Percent As Long, Cancel As Long) As eCompareResult
    'Callback proc - do not break or end
    On Error GoTo CancelOut
    frmTest.pgbProgress.Value = Percent
    If GetInput() Then
        DoEvents
        If bCancelFlag Then GoTo CancelOut
    End If
    CWP_Compare = StrComp(sA(ThisIdx), sA(ThanIdx), eCase)
    Exit Function
CancelOut:
    Cancel = True
End Function

'&&&&&&&&&&&&&& Demo code only &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Sub ResetCWPDemo(strA() As String)
    sA = strA
End Sub
Sub CWPMethod(ByVal eCaseA As VbCompareMethod)
    eCase = eCaseA
End Sub
Sub CWPCancel(ByVal bCancel As Boolean)
    bCancelFlag = bCancel
End Sub                               'attempt to make level performance test
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
