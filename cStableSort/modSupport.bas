Attribute VB_Name = "modSupport"
Option Explicit

Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

Private Declare Function PerfCount Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
Private Declare Function PerfFreq Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long

Private mCurFreq As Currency

Private Const NEG1 = -1&, n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&
Private Const n6 = 6&, n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&

Private Enum SAFEATURES
    FADF_AUTO = &H1               ' Array is allocated on the stack
    FADF_STATIC = &H2             ' Array is statically allocated
    FADF_EMBEDDED = &H4           ' Array is embedded in a structure
    FADF_FIXEDSIZE = &H10         ' Array may not be resized or reallocated
    FADF_BSTR = &H100             ' An array of BSTRs
    FADF_UNKNOWN = &H200          ' An array of IUnknown*
    FADF_DISPATCH = &H400         ' An array of IDispatch*
    FADF_VARIANT = &H800          ' An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8    ' Bits reserved for future use
    #If False Then
        Dim FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
    #End If
End Enum
Private Const VT_BYREF = &H4000&  ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix

' Used for unsigned arithmetic
Private Const DW_MSB = &H80000000 ' DWord Most Significant Bit

Private Type SAFEARRAY
    cDims       As Integer        ' Count of dimensions in this array
    fFeatures   As Integer        ' Bitfield flags indicating attributes of a particular array
    cbElements  As Long           ' Byte size of each element of the array
    cLocks      As Long           ' Number of times the array has been locked without corresponding unlock
    pvData      As Long           ' Pointer to the start of the array data (use only if cLocks > 0)
    cElements   As Long           ' Count of elements in this dimension
    lLbound     As Long           ' The lower-bounding index of this dimension
    lUbound     As Long           ' The upper-bounding index of this dimension
End Type

Public Enum eCompare
    Lesser = -1&
    Equal = 0&
    Greater = 1&
    #If False Then
        Dim Lesser, Equal, Greater
    #End If
End Enum

Public Enum eSortOrder
    Descending = -1&
    Default = 0&
    Ascending = 1&
    #If False Then
        Dim Descending, Default, Ascending
    #End If
End Enum

Public Enum eSortState
    Uninitialized = -1&
    Unsorted = 0&
    PreSorted = 1&
    PreRevSorted = 2&
    MostlySorted = 3&
    MostlyRevSorted = 4&
    SemiSorted = 5&
    SemiRevSorted = 6&
    #If False Then
        Dim Uninitialized, Unsorted, PreSorted, PreRevSorted, MostlySorted, MostlyRevSorted, SemiSorted, SemiRevSorted
    #End If
End Enum

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function ProfileStart() As Currency
    If mCurFreq = 0 Then PerfFreq mCurFreq
    If (mCurFreq) Then PerfCount ProfileStart
End Function

Public Function ProfileStop(ByVal curStart As Currency) As Currency
    If (mCurFreq) Then
        Dim curStop As Currency
        PerfCount curStop
        ProfileStop = (curStop - curStart) / mCurFreq ' cpu tick accurate
        curStop = 0
    End If
End Function

' + Inited Array ++++++++++++++++++++++++++++++++++++++++

' This function determines if the passed array is initialized,
' and if so will return -1.

' It will also optionally indicate whether the array can be redimmed;
' in which case it will return -2.

' If the array is uninitialized (never redimmed or has been erased)
' it will return 0 (zero).

Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant headache,
    ' and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY, lpSA As Long
    Dim iDataType As Integer, lOffset As Long
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, n2                       ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then           ' if a valid array was passed
        CopyMemByR lpSA, ByVal Sum(VarPtr(Arr), n8), n4 ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then                ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case lpSA already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMemByR lpSA, ByVal lpSA, n4             ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = (lpSA <> n0)
        If InitedArray Then
            CopyMemByR tSA.cDims, ByVal lpSA, n4
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + ((tSA.fFeatures And FADF_FIXEDSIZE) <> FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = n16 + ((tSA.cDims - n1) * n8)
            CopyMemByR tSA.cElements, ByVal Sum(lpSA, lOffset), n8
            tSA.lUbound = tSA.lLbound + tSA.cElements - n1
            If (lbA < tSA.lLbound) Then lbA = tSA.lLbound
            If (ubA > tSA.lUbound) Then ubA = tSA.lUbound
    End If: End If
UnInit:
End Function

' + Sum +++++++++++++++++++++++++++++++++++++++++++++++++

' Enables valid addition and subtraction of unsigned long ints.
' Treats lPtr as an unsigned long and returns an unsigned long.
' Allows safe arithmetic operations on memory address pointers.
' Assumes valid pointer and pointer offset.

Private Function Sum(ByVal lPtr As Long, ByVal lOffset As Long) As Long
    If lOffset > 0 Then
        If lPtr And DW_MSB Then ' if ptr < 0
           Sum = lPtr + lOffset ' ignors > unsigned int max
        ElseIf (lPtr Or DW_MSB) < -lOffset Then
           Sum = lPtr + lOffset ' result is below signed int max
        Else                    ' result wraps to min signed int
           Sum = (lPtr + DW_MSB) + (lOffset + DW_MSB)
        End If
    ElseIf lOffset = 0 Then
        Sum = lPtr
    Else 'If lOffset < 0 Then
        If (lPtr And DW_MSB) = 0 Then ' if ptr > 0
           Sum = lPtr + lOffset ' ignors unsigned int < zero
        ElseIf (lPtr - DW_MSB) >= -lOffset Then
           Sum = lPtr + lOffset ' result is above signed int min
        Else                    ' result wraps to max signed int
           Sum = (lOffset - DW_MSB) + (lPtr - DW_MSB)
        End If
    End If
End Function

Function strGetArraySortState(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal Method As VbCompareMethod, Optional ByVal FastDontTestMostly As Boolean) As eSortState '-©Rd 2005-
    Dim r As Long, d As Long, e As Long, cnt As Long
    Dim aeComp(n1 To n8) As eCompare
    Dim eComp As eCompare, bFlag As Boolean
    'On Error GoTo UnInit
    strGetArraySortState = Uninitialized
    If Not InitedArray(sA, lbA, ubA) Then Exit Function

    cnt = ubA - lbA
    If cnt < n1 Then strGetArraySortState = Unsorted: Exit Function
    r = cnt \ n12

    If r = n0 Then ' If small array
        ' Make 8 comparisons from lb to item 9
        cnt = n0
        For d = lbA To ubA - n1
            cnt = cnt + n1
            aeComp(cnt) = StrComp(sA(d), sA(d + n1), Method)
            If cnt = n8 Then Exit For
        Next
    Else 'If r > n0 Then
        ' Make 8 comparisons from 8% above lb to 72% step 8%
        d = lbA + r: e = d + r
        For cnt = n1 To n8
            aeComp(cnt) = StrComp(sA(d), sA(e), Method)
            d = e: e = e + r
        Next
        cnt = n8
    End If

    eComp = Greater ' Test sorted state
    Do: r = n0
        For d = n1 To cnt ' Not lo > hi (lo <= hi)...Not lo < hi (lo >= hi)
            r = r + Abs(Not aeComp(d) = eComp)
        Next                    '  descending          ascending
        e = (n2 + eComp) Mod n3 ' eComp=Lesser >> e=1|eComp=Greater >> e=0

        If r = n0 Then
            ' Previously sorted in other direction (with no equals)
            strGetArraySortState = PreRevSorted - e ' PreSorted|PreRevSorted
            eComp = -eComp

        ElseIf r = cnt - n1 Then
            ' May have been previously sorted
            strGetArraySortState = PreSorted + e + n4 ' SemiSorted|SemiRevSorted

        ElseIf r = cnt Then
            ' Previously sorted up to 72% (or 100% if i<8)
            strGetArraySortState = PreSorted + e ' PreSorted|PreRevSorted
        End If

        If strGetArraySortState = PreSorted Or strGetArraySortState = PreRevSorted Then
            If FastDontTestMostly Then
            ElseIf (cnt = n8) Then
                r = n0
                ' Compare the top 9 items of the array
                For d = ubA To lbA + n1 Step NEG1
                    r = r + n1
                    aeComp(r) = StrComp(sA(d - n1), sA(d), Method)
                    If r = n8 Then Exit For
                Next '-©Rd-
                For d = r To n1 Step NEG1
                    ' If any items are out of order
                    If aeComp(d) = eComp Then
                        strGetArraySortState = PreSorted + e + n2 ' MostlySorted|MostlyRevSorted
                        Exit For
                    End If
                Next
            End If
        ElseIf strGetArraySortState = Uninitialized Then
            If bFlag = False Then
                bFlag = True: eComp = Lesser
            Else
                ' This array is not sorted
                strGetArraySortState = Unsorted
        End If: End If

    Loop While strGetArraySortState = Uninitialized
UnInit:
End Function

Function strGetArraySortStateIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal Method As VbCompareMethod, Optional ByVal FastDontTestMostly As Boolean) As eSortState '-©Rd 2005-
    Dim r As Long, d As Long, e As Long, cnt As Long
    Dim aeComp(n1 To n8) As eCompare
    Dim eComp As eCompare, bFlag As Boolean
    'On Error GoTo UnInit
    strGetArraySortStateIndexed = Uninitialized
    If InitedArray(sA, lbA, ubA) = False Then Exit Function

    cnt = ubA - lbA
    If cnt < n1 Then strGetArraySortStateIndexed = Unsorted: Exit Function
    r = cnt \ n12

    If r = n0 Then ' If small array
        ' Make 8 comparisons from lb to item 9
        cnt = n0
        For d = lbA To ubA - n1
            cnt = cnt + n1
            aeComp(cnt) = StrComp(sA(idxA(d)), sA(idxA(d + n1)), Method)
            If cnt = n8 Then Exit For
        Next
    Else 'If r > n0 Then
        ' Make 8 comparisons from 8% above lb to 72% step 8%
        d = lbA + r: e = d + r
        For cnt = n1 To n8
            aeComp(cnt) = StrComp(sA(idxA(d)), sA(idxA(e)), Method)
            d = e: e = e + r
        Next
        cnt = n8
    End If

    eComp = Greater ' Test sorted state
    Do: r = n0
        For d = n1 To cnt ' Not lo > hi (lo <= hi)...Not lo < hi (lo >= hi)
            r = r + Abs(Not aeComp(d) = eComp)
        Next                    '  descending          ascending
        e = (n2 + eComp) Mod n3 ' eComp=Lesser >> e=1|eComp=Greater >> e=0

        If r = n0 Then
            ' Previously sorted in other direction (with no equals)
            strGetArraySortStateIndexed = PreRevSorted - e ' PreSorted|PreRevSorted
            eComp = -eComp

        ElseIf r = cnt - n1 Then
            ' May have been previously sorted
            strGetArraySortStateIndexed = PreSorted + e + n4 ' SemiSorted|SemiRevSorted

        ElseIf r = cnt Then
            ' Previously sorted up to 72% (or 100% if i<8)
            strGetArraySortStateIndexed = PreSorted + e ' PreSorted|PreRevSorted
        End If

        If strGetArraySortStateIndexed = PreSorted Or strGetArraySortStateIndexed = PreRevSorted Then
            If FastDontTestMostly Then
            ElseIf (cnt = n8) Then
                r = n0
                ' Compare the top 9 items of the array
                For d = ubA To lbA + n1 Step NEG1
                    r = r + n1
                    aeComp(r) = StrComp(sA(idxA(d - n1)), sA(idxA(d)), Method)
                    If r = n8 Then Exit For
                Next '-©Rd-
                For d = r To n1 Step NEG1
                    ' If any items are out of order
                    If aeComp(d) = eComp Then
                        strGetArraySortStateIndexed = PreSorted + e + n2 ' MostlySorted|MostlyRevSorted
                        Exit For
                    End If
                Next
            End If
        ElseIf strGetArraySortStateIndexed = Uninitialized Then
            If bFlag = False Then
                bFlag = True: eComp = Lesser
            Else
                ' This array is not sorted
                strGetArraySortStateIndexed = Unsorted
        End If: End If

    Loop While strGetArraySortStateIndexed = Uninitialized
UnInit:
End Function

Function strVerifyIndexed(sA() As String, lA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal Method As VbCompareMethod, Optional ByVal Order As eSortOrder = Ascending) As Boolean
    On Error GoTo FreakOut
    Dim walk As Long
    For walk = lbA + n1 To ubA
        If StrComp(sA(lA(walk - n1)), sA(lA(walk)), Method) = Order Then Exit Function
    Next
FreakOut:
    strVerifyIndexed = (walk > ubA)
End Function

Public Sub GetFileUBs()
    Dim s As String
    s = Dir$(App.Path & "\ArrayFiles\Arr*.dat")
    Do Until LenB(s) = n0
        AddCombo Mid$(s, 4, InStr(s, "_") - 4)
        s = Dir$
    Loop
End Sub

Public Sub AddCombo(ByVal Max As Long)
    Dim idx As Integer
    With frmTest.cboUB
        For idx = 0 To .ListCount - 1
            ' If cboUB.Text already exists exit sub
            If (.List(idx) = CStr(Max)) Then Exit Sub
            ' If found the correct index exit for
            If (Val(.List(idx)) > Max) Then Exit For
        Next
        .AddItem Max, idx
    End With
End Sub

Public Sub ResetLabels(Optional ByVal Index As Long = 0)
    Dim x As Long
    With frmTest
        If Index > 0 Then
            .lblResults(Index) = "9999999"
        Else
            For x = 1 To .lblResults.Count
                .lblResults(x) = "9999999"
            Next x
        End If
    End With
End Sub

' Find the True option from a control array of OptionButtons
Public Function GetOption(opts As Object) As Integer
    ' Assume no option set True
    GetOption = -1
    On Error GoTo GetOptionFail
    Dim opt As OptionButton
    For Each opt In opts
        If opt.Value Then
            GetOption = opt.Index
            Exit Function
        End If
    Next
GetOptionFail:
End Function
