VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStableSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++©Rd+

' Generic stable indexed data sorting class.

' You are free to use any part or all of this code even for
' commercial purposes in any way you wish under the one condition
' that no copyright notice is moved or removed from where it is.

' This sorting class has the following features and limitations:
' - It requires an index (long) array to be passed instead of the
'   source array.
' - Because it uses an index array to sort the items no 'SwapItem'
'   routine is required and so is considerably faster sorting the
'   long array items internally. This permits smarter sorting that
'   is not limited to the item swap technique, which allows for
'   fast copymemory operations that can manipulate multiple items
'   simultaneously.
' - The index array is automatically initialized if it is passed
'   uninitialized or erased, or if the bounds passed are invalid
'   for the index arrays current boundaries.
' - The index array can be passed again without erasing it. See
'   the header comments for PrepareIndexArray for more details.
' - It can handle zero or higher based source arrays.
' - It can handle negative lb and positive ub.
' - It can handle negative lb and zero or negative ub.
' - It cannot handle lb greater than ub.
' - It does not validate the array bounds passed.
' - You must write code to handle the Compare event, which must
'   assign to the Result parameter an eComp enumeration value.
' - You can specify which dimension to sort in the Compare event.
' - It does not validate the specified dimension.
' - You can specify the sort order (ascending|descending).
' - You can Cancel sorting within the Compare event.
' - A progress parameter in the Compare event indicates the
'   percentage completed as a long value of 1 to 100.

' For comments, suggestions or bug reports you can
' contact me at rd•edwards•bigpond•com.

' + Usage +++++++++++++++++++++++++++++++++++++++++++++++++

' To use this class declare it withevents in a form or class:

'Private WithEvents cSort as cStableSort

' Before using it you must create it:

'Set cSort = New cStableSort

' Optionally to prepare the index array before sorting
' for the first time:

'ReDim idxA(lb To ub) As Long
'For i = lb To ub
'    idxA(i) = i
'Next

' Preparing to sort:

'Private mDim As Long
'Property Get Dimension() As Long
'    If mDim < n1 Then mDim = n1
'    Dimension = mDim
'End Property
'
'Property Let Dimension(ByVal NewDimension As Long)
'    If NewDimension < n1 Then
'        mDim = n1
'    Else
'        mDim = NewDimension
'    End If
'End Property

'Private Sub cmdCancel_Click()
'    mCancelFlag = True
'End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Sorting your array:

'pgbProgress.min = 1
'pgbProgress.max = 100
'cSort.Order = Ascend
'mCancelFlag = False
'cSort.Sort idxA, lb, ub

'Private Sub cSort_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eComp, ByVal Percent As Long, Cancel As Boolean)
'    DoEvents
'    If mCancelFlag Then Cancel = True: Exit Sub
'    pgbProgress.Value = Percent
'    If srcA(mDim, ThisIdx) > srcA(mDim, ThanIdx) Then
'        Result = Greater
'    ElseIf srcA(mDim, ThisIdx) < srcA(mDim, ThanIdx) Then
'        Result = Lesser
'    Else
'        Result = Equal
'    End If
'End Sub

' Notice you can safely cancel at the start of the Compare event
' by setting Cancel = True and immediately exiting if you wish to
' avoid executing code in the Compare event when the user cancels.

' To access the array items in sorted order:

'For i = lb To ub
'    item = srcA(mDim, idxA(i))
'Next

'Set cSorter = Nothing

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Event Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eComp, ByVal Percent As Long, Cancel As Boolean)

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' An object that raises events is called an event source.
' To handle the events raised by an event source, you can declare
' a variable of the object's class using the WithEvents keyword.
'
' The variable is declared in the Declarations section because
' WithEvents variables must be module-level variables.
' This is true regardless of the type of module you place them in.
'
' A WithEvents variable cannot be a generic object variable.
' That is, you cannot declare it As Object — you must specify the
' class name when you declare the variable.
'
' You cannot declare a WithEvents variable As New.
' The event source object must be explicitly created and assigned
' to the WithEvents variable.
'
' You cannot declare WithEvents variables in a standard module.
' You can declare them only in class modules, form modules, and
' other modules that define classes.
'
' When you declare a variable WithEvents at design time, there is
' no object associated with it. A WithEvents variable is just like
' any other object variable. You have to create an object and assign
' a reference to the object to the WithEvents variable.
'
' Important - A WithEvents variable contains an object reference, just
' like any other object variable. This object reference counts toward
' keeping the object alive. When you are setting all references to an
' object to Nothing in order to destroy it, don't forget the variables
' you declared WithEvents.

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Declare some CopyMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

' More efficient repeated use of numeric literals
Private Const n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&, n6 = 6&
Private Const n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&, n64 = 64&
Private Const n10K As Long = 10000&
Private Const n20K As Long = 20000&
Private Const n50K As Long = 50000

' Used for unsigned arithmetic
Private Const DW_MSB = &H80000000 ' DWord Most Significant Bit

Private Const rRunner4 As Single = 0.0025 '0.002<<reverse-sorting-0.003-unsorted>>0.004
Private Const rRunner5 As Single = 0.0015 '0.001<<reverse-sorting-unsorted>>0.002

Private ss2Lb() As Long, ss2Ub() As Long ' Stable non-recursive quicksort stacks
Private tw4Lb() As Long, tw4Ub() As Long ' Stable insert/binary runner stacks
Private lA_1() As Long, lA_2() As Long   ' Stable quicksort working buffers
Private ss2Max As Long, tw4Max As Long, bufMax As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Enum eComp
    Lesser = -1&
    Equal = 0&
    Greater = 1&
    #If False Then
        Dim Lesser, Equal, Greater
    #End If
End Enum

Public Enum eDir
    Descending = -1&
    Default = 0&
    Ascending = 1&
    #If False Then
        Dim Descending, Default, Ascending
    #End If
End Enum

Private Const mDef_Dir As Long = Ascending
Private mComp As eDir

Private bBuzy As Boolean

' + Order Property ++++++++++++++++++++++++++++++++++++++

' This property specifies whether the sorting direction is
' in ascending or descending order.

Property Get Order() As eDir
    If mComp = Default Then mComp = mDef_Dir
    Order = mComp
End Property

Property Let Order(ByVal eNewDirection As eDir)
    If bBuzy Then Beep: Exit Property
    If eNewDirection = Default Then
        If mComp = Default Then mComp = mDef_Dir
    Else
        mComp = eNewDirection
    End If
End Property

' + Indexed Stable QuickSort ++++++++++++++++++++++++++++++++

' Indexed stable non-recursive quicksort based algorithm which
' raises a Compare Event but no 'SwapItem' required.

Sub Sort(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long) '-©Rd 2006-
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, run As Long, cast As Long
    Dim idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim ceil As Long, mezz As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim idxItem As Long, lpLB As Long
    Dim lPercent As Long, lpI As Long
    Dim eResult As eComp, lProg As Long
    Dim eComp As eDir
    Dim bCancel As Boolean
    Dim rTotal As Single

    If bBuzy Then Beep: Exit Sub
    rTotal = ubA - lbA + n1              ' Grab array item count
    If (rTotal < n2) Then Exit Sub       ' If nothing to do then exit
    eComp = Order                        ' Initialize compare variable
    pvt = (rTotal \ n64) + n32           ' Allow for worst case senario + some
    bBuzy = True

    ValidateIndexArray idxA, lbA, ubA           ' Validate the index array
    InitializeStacks ss2Lb, ss2Ub, ss2Max, pvt  ' Initialize pending boundary stacks
    InitializeStacks tw4Lb, tw4Ub, tw4Max, pvt  ' Initialize pending runner stacks
    InitializeStacks lA_1, lA_2, bufMax, rTotal ' Initialize working buffers

    rTotal = rTotal * 0.66
    lpL_1 = VarPtr(lA_1(n0))                  ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))                  ' Cache pointer to upper buffer
    lpI = Sum(VarPtr(idxA(lbA)), -(lbA * n4)) ' Cache pointer to the index array

    Do: lProg = lProg + n1
        lPercent = CLng(100 * lProg / rTotal)
        If lPercent < n1 Then
            lPercent = n1
        ElseIf lPercent > 100 Then       ' Progress Indicator
            lPercent = 100
        End If
        ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA   ' Get pivot index position
        idxItem = idxA(pvt)              ' Grab current index into item

        For idx = lbA To pvt - n1
            RaiseEvent Compare(idxA(idx), idxItem, eResult, lPercent, bCancel)
            If bCancel Then GoTo ExitSub
            If (eResult = eComp) Then    ' (idx > item)
                lA_2(ptr2) = idxA(idx)   ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(idx)   ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            RaiseEvent Compare(idxItem, idxA(idx), eResult, lPercent, bCancel)
            If bCancel Then GoTo ExitSub
            If (eResult = eComp) Then    ' (idx < item)
                lA_1(ptr1) = idxA(idx)   ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(idx)   ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2i ©Rd-
        lpLB = VarPtr(idxA(lbA))         ' Cache pointer to current lb
        CopyMemByV lpLB, lpL_1, ptr1 * n4
        idxA(lbA + ptr1) = idxItem       ' 1 2 item 3 4
        CopyMemByV Sum(lpLB, (ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If (ubA - lbA < n64) Then        ' Ignore false indicators
            If (inter2 = ptr2) Then      ' Reverse indicator
            ElseIf (ubA - lbA < n3) Then ' Delegate to built-in Repeater on tiny chunks
                For walk = lbA + n1 To ubA
                    lProg = lProg + n1
                    lPercent = CLng(100 * lProg / rTotal)
                    If lPercent < n1 Then
                        lPercent = n1
                    ElseIf lPercent > 100 Then                  ' Progress Indicator
                        lPercent = 100
                    End If
                    idxItem = idxA(walk)     ' Grab current value
                    find = walk
                    RaiseEvent Compare(idxA(find - n1), idxItem, eResult, lPercent, bCancel)
                    If bCancel Then GoTo ExitSub
                    Do While eResult = eComp
                        find = find - n1
                        If (find = lbA) Then Exit Do
                        RaiseEvent Compare(idxA(find - n1), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                    Loop '-Repeater v45i ©Rd-
                    If (find < walk) Then    ' Move items up 1, sA(find) = item
                        CopyMemByV Sum(lpI, (find + n1) * n4), Sum(lpI, find * n4), (walk - find) * n4
                        idxA(find) = idxItem ' Re-assign current item index into found pos
                End If: Next
                ptr1 = n0: ptr2 = n0
            End If
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then      ' Reverse
            ElseIf (ptr1 = n0) Then      ' Reverse Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                  ' Runners dislike super large ranges
                    idxItem = idxA(lbA + ptr1 - n1)
                    optimal = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > optimal                                   ' Runner do loop
                        RaiseEvent Compare(idxA(run - n1), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                        If Not eResult = eComp Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, run * n4), cast * n4      ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpI, run * n4), Sum(lpI, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpI, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    idxItem = idxA(base)
                    pvt = lbA + ptr1 + inter2
                    optimal = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < optimal                                   ' Runner do loop
                        RaiseEvent Compare(idxA(run + n1), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                        If Not eResult = eComp Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpI, (base + cast) * n4), Sum(lpI, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpI, base * n4), lpL_1, cast * n4     ' Re-assign items into position immediately above current item
            End If: End If: End If
        ElseIf (inter2 = n0) Then
            If (inter1 = ptr1) Then       ' Refresh
                ' Delegate to built-in Insert/Binary hybrid on ideal data state
                walk = lbA: mezz = ubA: idx = n0                                  ' Initialize our walker variables
                optimal = GetOptimalRange(ubA - lbA + n1)                         ' Get runners optimal range
                If optimal > tw4Max Then                                          ' Ensure enough stack space
                    InitializeStacks tw4Lb, tw4Ub, tw4Max, optimal
                End If
                Do While walk < mezz ' ----==================================---- ' Do the twist while there's more items
                    walk = walk + n1                                              ' Walk up the array and use binary search to insert each item down into the sorted lower array
                    lProg = lProg + n1
                    lPercent = CLng(100 * lProg / rTotal)
                    If lPercent < n1 Then
                        lPercent = n1
                    ElseIf lPercent > 100 Then                                    ' Progress Indicator
                        lPercent = 100
                    End If
                    idxItem = idxA(walk)                                          ' Grab current index
                    find = walk                                                   ' Default to current position
                    ceil = walk - n1                                              ' Set ceiling to current position - 1
                    base = lbA                                                    ' Set base to lower bound
                    RaiseEvent Compare(idxA(ceil), idxItem, eResult, lPercent, bCancel)
                    If bCancel Then GoTo ExitSub
                    Do While eResult = eComp '                                '   ' While current item must move down
                        midd = (base + ceil) \ n2                                 ' Find mid point
                        RaiseEvent Compare(idxA(midd), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                        Do Until eResult = eComp                                  ' Step back up if below
                            base = midd + n1                                      ' Bring up the base
                            midd = (base + ceil) \ n2                             ' Find mid point
                            If midd = ceil Then Exit Do                           ' If we're up to ceiling
                            RaiseEvent Compare(idxA(midd), idxItem, eResult, lPercent, bCancel)
                            If bCancel Then GoTo ExitSub
                        Loop                                                      ' Out of loop >= target pos
                        find = midd                                               ' Set provisional to new ceiling
                        If find = base Then Exit Do                               ' If we're down to base
                        ceil = midd - n1                                          ' Bring down the ceiling
                        RaiseEvent Compare(idxA(ceil), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                    Loop '-Twister v4i ©Rd-   .       . ...   .               .   ' Out of binary search loops
                    If (find < walk) Then                                         ' If current item needs to move down
                        idxItem = idxA(find)                                      ' Grab found index
                        run = walk + n1
                        Do Until run > mezz Or run - walk > optimal               ' Runner do loop
                            RaiseEvent Compare(idxItem, idxA(run), eResult, lPercent, bCancel)
                            If bCancel Then GoTo ExitSub
                            If Not eResult = eComp Then Exit Do
                            run = run + 1
                        Loop: cast = run - walk
                        CopyMemByV lpL_2, Sum(lpI, walk * n4), cast * n4          ' Grab current value(s)
                        CopyMemByV Sum(lpI, (find + cast) * n4), Sum(lpI, find * n4), (walk - find) * n4 ' Move up items
                        CopyMemByV Sum(lpI, find * n4), lpL_2, cast * n4          ' Re-assign current value(s) into found pos
                        If cast > n1 Then
                            If run <= mezz Then
                                idx = idx + n1
                                tw4Lb(idx) = run - n1
                                tw4Ub(idx) = mezz
                            End If
                            walk = find
                            mezz = find + cast - n1
                    End If: End If
                    If walk = mezz Then
                        If idx Then
                            walk = tw4Lb(idx)
                            mezz = tw4Ub(idx)
                            idx = idx - n1
                End If: End If: Loop         ' Out of walker do loop
                ' ----===========================----
                ptr1 = n0: ptr2 = n0
            ElseIf (ptr2 = n0) Then     ' Pretty
                If (ptr1 > inter1) And (inter1 < n50K) Then                  ' Runners dislike super large ranges
                    idxItem = idxA(lbA + ptr1 - n1)
                    optimal = lbA + (inter1 \ n2)
                    run = lbA + inter1
                    Do While run > optimal                                   ' Runner do loop
                        RaiseEvent Compare(idxA(run - n1), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                        If Not eResult = eComp Then Exit Do
                        run = run - n1
                    Loop: cast = lbA + inter1 - run
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, run * n4), cast * n4      ' Grab items that stayed below current that should also be above items that have moved down below current
                        CopyMemByV Sum(lpI, run * n4), Sum(lpI, (lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                        CopyMemByV Sum(lpI, (lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                    End If
                End If ' 1 2 1r item 4r 3 4
                If (inter2) And (ptr2 - inter2 < n50K) Then
                    base = lbA + ptr1 + n1
                    idxItem = idxA(base)
                    pvt = lbA + ptr1 + inter2
                    optimal = pvt + ((ptr2 - inter2) \ n2)
                    run = pvt
                    Do While run < optimal                                   ' Runner do loop
                        RaiseEvent Compare(idxA(run + n1), idxItem, eResult, lPercent, bCancel)
                        If bCancel Then GoTo ExitSub
                        If Not eResult = eComp Then Exit Do
                        run = run + n1
                    Loop: cast = run - pvt
                    If cast Then
                        CopyMemByV lpL_1, Sum(lpI, (pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                        CopyMemByV Sum(lpI, (base + cast) * n4), Sum(lpI, base * n4), inter2 * n4 ' Move up items
                        CopyMemByV Sum(lpI, base * n4), lpL_1, cast * n4     ' Re-assign items into position immediately above current item
        End If: End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: ss2Lb(cnt) = lbA + ptr1 + n1: ss2Ub(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = ss2Lb(cnt): ubA = ss2Ub(cnt): cnt = cnt - n1
    End If: Loop
ExitSub:
    bBuzy = False
End Sub

' + Validate Index Array +++++++++++++++++++++++++++++++++++++

' This will prepare the passed index array if it is not already.

' This sub-routine determines if the index array passed is either:
' [A] uninitialized or Erased
'     initialized to invalid boundaries
'     initialized to valid boundaries but not prepared
' [B] initialized to extended boundaries and not fully prepared
' [C] prepared for the sort process by the For loop
'     has been modified by a previous sort process

' If the condition is determined to be [A] then it is prepared by
' executing the For loop code, if the condition is determined to
' be [B] then it is prepared only from the old ub to the new ub,
' otherwise if [C] nothing is done.

' This permits subsequent sorting of the data without interfering
' with the index array if it is already sorted (based on criteria
' that may differ from the current process, for example, or some
' items have been modified in the sorted array).

' It also permits refresh-sorting of data that has additional
' items added to the top of the sorted array without interfering
' with the index array and so does not require a full resort.

' Otherwise, it ensures that the index array is in the required
' pre-sort state produced by the For loop.

Sub ValidateIndexArray(idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim bReDim As Boolean, bReDimPres As Boolean
    Dim lb As Long, ub As Long, j As Long
    lb = &H80000000: ub = &H7FFFFFFF
    bReDim = Not InitedArray(idxA, lb, ub)
    If bReDim = False Then
        bReDim = (lbA < lb)
        bReDimPres = (ubA > ub)
    End If '-©Rd-
    If bReDim Then
        ReDim idxA(lbA To ubA) As Long
    ElseIf bReDimPres Then
        ReDim Preserve idxA(lb To ubA) As Long
    End If
    If (idxA(ubA) = n0) Then
        If (idxA(lbA) = n0) Then
            For j = lbA To ubA
                idxA(j) = j
            Next
        ElseIf bReDimPres Then
            For j = ub + n1 To ubA
                idxA(j) = j
            Next
        End If
    End If
End Sub

' + Sum +++++++++++++++++++++++++++++++++++++++++++++++++

' Enables valid addition and subtraction of unsigned long ints.
' Treats lPtr as an unsigned long and returns an unsigned long.
' Allows safe arithmetic operations on memory address pointers.
' Assumes valid pointer and pointer offset.

Private Function Sum(ByVal lPtr As Long, ByVal lOffset As Long) As Long
    If lOffset > n0 Then
        If lPtr And DW_MSB Then ' if ptr < 0
           Sum = lPtr + lOffset ' ignors > unsigned int max
        ElseIf (lPtr Or DW_MSB) < -lOffset Then
           Sum = lPtr + lOffset ' result is below signed int max
        Else                    ' result wraps to min signed int
           Sum = (lPtr + DW_MSB) + (lOffset + DW_MSB)
        End If
    ElseIf lOffset = n0 Then
        Sum = lPtr
    Else 'If lOffset < 0 Then
        If (lPtr And DW_MSB) = n0 Then ' if ptr > 0
           Sum = lPtr + lOffset ' ignors unsigned int < zero
        ElseIf (lPtr - DW_MSB) >= -lOffset Then
           Sum = lPtr + lOffset ' result is above signed int min
        Else                    ' result wraps to max signed int
           Sum = (lOffset - DW_MSB) + (lPtr - DW_MSB)
        End If
    End If
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Function GetOptimalRange(ByVal lCount As Long, Optional ByVal fOptimal As Boolean) As Long
    Dim optimal As Long, range As Single ' CraZy performance curve
    If lCount > n20K Then optimal = n12 * (lCount \ n10K - n2)
    If fOptimal Then range = rRunner5 Else range = rRunner4
    GetOptimalRange = (lCount * range) - optimal + n4
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub InitializeStacks(LBstack() As Long, UBstack() As Long, ByRef pCurMax As Long, ByVal NewMax As Long)
    If NewMax > pCurMax Then
        ReDim LBstack(n0 To NewMax) As Long   ' Stack to hold pending lower boundries
        ReDim UBstack(n0 To NewMax) As Long   ' Stack to hold pending upper boundries
        pCurMax = NewMax
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!
