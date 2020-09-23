Attribute VB_Name = "CompareRoutines"
Option Explicit

Function CompareName(First As CEmployee, Second As CEmployee, unused1 As Long, unused2 As Long) As eCompareResult
    CompareName = StrComp(First.ReversedName, Second.ReversedName, vbTextCompare)
End Function

Function CompareDeptName(First As CEmployee, Second As CEmployee, unused1 As Long, unused2 As Long) As eCompareResult
    Dim result As Integer
    result = StrComp(First.Dept, Second.Dept, vbTextCompare)
    If result = 0 Then
        result = StrComp(First.ReversedName, Second.ReversedName, vbTextCompare)
    End If
    CompareDeptName = result
End Function

Function CompareSalaryName(First As CEmployee, Second As CEmployee, unused1 As Long, unused2 As Long) As eCompareResult
    Dim result As Integer
    result = -Sgn(First.Salary - Second.Salary)
    If result = 0 Then
        result = StrComp(First.ReversedName, Second.ReversedName, vbTextCompare)
    End If
    CompareSalaryName = result
End Function

Function CompareDeptSalaryName(First As CEmployee, Second As CEmployee, unused1 As Long, unused2 As Long) As eCompareResult
    Dim result As Integer
    result = StrComp(First.Dept, Second.Dept, vbTextCompare)
    If result = 0 Then
        result = -Sgn(First.Salary - Second.Salary)
        If result = 0 Then
            result = StrComp(First.ReversedName, Second.ReversedName, vbTextCompare)
        End If
    End If
    CompareDeptSalaryName = result
End Function

Function CompareDates(First As CEmployee, Second As CEmployee, unused1 As Long, unused2 As Long) As eCompareResult
    If First.Hired > Second.Hired Then
        CompareDates = crGreater
    ElseIf Second.Hired > First.Hired Then
        CompareDates = crLess
    End If
End Function

Function CompareLong(Long1 As Long, Long2 As Long, unused1 As Long, unused2 As Long) As eCompareResult
    If Long1 > Long2 Then
        CompareLong = crGreater
    ElseIf Long2 > Long1 Then
        CompareLong = crLess
    End If
End Function
