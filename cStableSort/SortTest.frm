VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTest 
   Caption         =   " Generic Stable Indexed Data Sorting Classes"
   ClientHeight    =   6990
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   8700
   Icon            =   "SortTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDynamic 
      Caption         =   "Dynamic Test Data State"
      Height          =   465
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "Test results are re-sorted"
      Top             =   4650
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cStableSortCWP.Sort"
      Height          =   330
      Left            =   120
      TabIndex        =   17
      Top             =   4245
      Width           =   1900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel sorting process now !"
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3780
      Width           =   2895
   End
   Begin VB.CommandButton cmdUDTs 
      Caption         =   "UDT arrays, multi-dimensional arrays, collections, variant arrays, etc"
      Height          =   585
      Left            =   150
      TabIndex        =   24
      Top             =   5880
      Width           =   2880
   End
   Begin VB.CommandButton cmdSortClassInfo 
      Caption         =   "About the sort classes..."
      Height          =   330
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   2625
   End
   Begin VB.CommandButton cmdResorterHlp 
      Caption         =   "More about the resorter..."
      Height          =   330
      Left            =   240
      TabIndex        =   23
      Top             =   5430
      Width           =   2625
   End
   Begin VB.CheckBox chkCaseSens 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "What the.."
      Height          =   330
      Left            =   1980
      TabIndex        =   8
      Top             =   1470
      Width           =   975
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      ItemData        =   "SortTest.frx":0442
      Left            =   300
      List            =   "SortTest.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1910
      Width           =   1125
   End
   Begin VB.TextBox txtLB 
      Height          =   285
      Left            =   780
      TabIndex        =   2
      Text            =   "5"
      Top             =   660
      Width           =   585
   End
   Begin VB.CommandButton cmdWriteU 
      Caption         =   "write file (original)"
      Height          =   330
      Left            =   330
      TabIndex        =   7
      Top             =   1470
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel sorting process now !"
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2850
      Width           =   2895
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Labels"
      Height          =   330
      Left            =   1830
      TabIndex        =   20
      Top             =   4740
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cStableSorter.Sort"
      Height          =   330
      Left            =   120
      TabIndex        =   14
      Top             =   3330
      Width           =   1900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   210
      Width           =   5280
   End
   Begin VB.CheckBox chkDirection 
      Caption         =   "Descending"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   5160
      Width           =   1185
   End
   Begin VB.ComboBox cboUB 
      Height          =   315
      ItemData        =   "SortTest.frx":0462
      Left            =   1710
      List            =   "SortTest.frx":0464
      TabIndex        =   4
      Text            =   "17275"
      Top             =   660
      Width           =   945
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "read file"
      Height          =   330
      Left            =   1710
      TabIndex        =   6
      Top             =   1050
      Width           =   915
   End
   Begin VB.CommandButton cmdWriteS 
      Caption         =   "write file (sorted)"
      Height          =   330
      Left            =   1470
      TabIndex        =   10
      Top             =   1890
      Width           =   1305
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "build array"
      Height          =   330
      Left            =   570
      TabIndex        =   5
      Top             =   1050
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cStableSort.Sort"
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   2370
      Width           =   1900
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   210
      Left            =   180
      TabIndex        =   26
      Top             =   6660
      Width           =   8260
      _ExtentX        =   14579
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
      Max             =   500
   End
   Begin VB.Label lblResults 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   2070
      TabIndex        =   18
      Top             =   4275
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "LB:"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   705
      Width           =   435
   End
   Begin VB.Label lblResults 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2070
      TabIndex        =   15
      Top             =   3345
      Width           =   945
   End
   Begin VB.Label lblResults 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2070
      TabIndex        =   12
      Top             =   2385
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "UB:"
      Height          =   225
      Index           =   1
      Left            =   1260
      TabIndex        =   3
      Top             =   705
      Width           =   435
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetInput Lib "user32" Alias "GetInputState" () As Long

Private lMin As Long
Private lMax As Long

Private suffix As String
Private lPrevChkState As Long
Private eDirect As eSortOrder
Private eCase As VbCompareMethod

Private aIdx() As Long 'Index to elements in sA
Private sA() As String

Private WithEvents cSSort As cStableSort
Attribute cSSort.VB_VarHelpID = -1

Private cSSortCWP As cStableSortCWP

Private cSSorter As cStableSorter
Implements ICompareClient

Private bCancelFlag As Boolean
Private iLongAlign As Integer

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cSSort_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eComp, ByVal Percent As Long, Cancel As Boolean)
    pgbProgress.Value = Percent
    Result = StrComp(sA(ThisIdx), sA(ThanIdx), eCase)
    If GetInput() Then
        DoEvents
        If bCancelFlag Then Cancel = True
    End If
End Sub

Private Sub ICompareClient_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eiCompare, ByVal Percent As Long, Cancel As Boolean)
    pgbProgress.Value = Percent
    Result = StrComp(sA(ThisIdx), sA(ThanIdx), eCase)
    If GetInput() Then
        DoEvents
        If bCancelFlag Then Cancel = True
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdCancel_Click(Index As Integer)
    bCancelFlag = True
    Call CWPCancel(True)
End Sub

Private Sub Command1_Click()
    Dim curElapse As Currency, ss$, r1 As Single
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    DizEnableEm
    ss = "cStableSort.Sort..." & lMax & vbNewLine
    If suffix = "_S" Or chkDynamic <> 0 Then
        ss = ss & "pre-sorted" & vbNewLine
    Else
        ss = ss & "un-sorted" & vbNewLine
    End If
    ss = ss & "Desc = " & CBool(chkDirection.Value) & vbNewLine

    If chkDynamic = 0 Then ResetIdxArray

    With cSSort
        .Order = eDirect
        pgbProgress.Min = 1
        pgbProgress.Max = 100

        bCancelFlag = False

        curElapse = ProfileStart
        .Sort aIdx, lMin, lMax

    End With

    If (curElapse > 0) Then
        r1 = CSng(ProfileStop(curElapse))
        If r1 < CSng(lblResults(1)) Then lblResults(1) = Format$(r1, "##0.0000")
        lblResults(1).Refresh
        ' Display it in the text box

ss = ss & "Sort verification " & IIf(strVerifyIndexed(sA, aIdx, lMin, lMax, eCase, eDirect), "successful", "FAILED!") & vbNewLine

        ss = ss & Format$(r1, "##0.0000") & " seconds! " & vbNewLine & vbNewLine
        Text1.SelStart = Len(Text1)
        Text1.SelText = ss   '& sa(lMin) & " " & sa(lMax)
    End If
    Screen.MousePointer = vbDefault
    DizEnableEm True
    pgbProgress.Value = 1
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    DizEnableEm True
    MsgBox "Command1_Click in frmTest Error - " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command2_Click()
    Dim curElapse As Currency, ss$, r1 As Single
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    DizEnableEm
    ss = "cStableSorter.Sort..." & lMax & vbNewLine
    If suffix = "_S" Or chkDynamic <> 0 Then
        ss = ss & "pre-sorted" & vbNewLine
    Else
        ss = ss & "un-sorted" & vbNewLine
    End If
    ss = ss & "Desc = " & CBool(chkDirection.Value) & vbNewLine

    If chkDynamic = 0 Then ResetIdxArray

    If cSSorter Is Nothing Then Exit Sub
    With cSSorter
        .Order = eDirect
        pgbProgress.Min = 1
        pgbProgress.Max = 100

        bCancelFlag = False

        curElapse = ProfileStart
        .Sort aIdx, lMin, lMax
        
    End With

    If (curElapse > 0) Then
        r1 = CSng(ProfileStop(curElapse))
        If r1 < CSng(lblResults(2)) Then lblResults(2) = Format$(r1, "##0.0000")
        lblResults(2).Refresh
        ' Display it in the text box

ss = ss & "Sort verification " & IIf(strVerifyIndexed(sA, aIdx, lMin, lMax, eCase, eDirect), "successful", "FAILED!") & vbNewLine

        ss = ss & Format$(r1, "##0.0000") & " seconds! " & vbNewLine & vbNewLine
        Text1.SelStart = Len(Text1)
        Text1.SelText = ss   '& sa(lMin) & " " & sa(lMax)
    End If
    Screen.MousePointer = vbDefault
    DizEnableEm True
    pgbProgress.Value = 1
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    DizEnableEm True
    MsgBox "Command2_Click in frmTest Error - " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command3_Click()
    Dim curElapse As Currency, ss$, r1 As Single
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    DizEnableEm
    ss = "cStableSortCWP.Sort..." & lMax & vbNewLine
    If suffix = "_S" Or chkDynamic <> 0 Then
        ss = ss & "pre-sorted" & vbNewLine
    Else
        ss = ss & "un-sorted" & vbNewLine
    End If
    ss = ss & "Desc = " & CBool(chkDirection.Value) & vbNewLine

    If chkDynamic = 0 Then ResetIdxArray

    If cSSortCWP Is Nothing Then Exit Sub
    With cSSortCWP
        .Order = eDirect
        pgbProgress.Min = 1
        pgbProgress.Max = 100

        bCancelFlag = False

        curElapse = ProfileStart
        .Sort aIdx, lMin, lMax, AddressOf CWP_Compare

    End With

    If (curElapse > 0) Then
        r1 = CSng(ProfileStop(curElapse))
        If r1 < CSng(lblResults(3)) Then lblResults(3) = Format$(r1, "##0.0000")
        lblResults(3).Refresh
        ' Display it in the text box

ss = ss & "Sort verification " & IIf(strVerifyIndexed(sA, aIdx, lMin, lMax, eCase, eDirect), "successful", "FAILED!") & vbNewLine

        ss = ss & Format$(r1, "##0.0000") & " seconds! " & vbNewLine & vbNewLine
        Text1.SelStart = Len(Text1)
        Text1.SelText = ss   '& sa(lMin) & " " & sa(lMax)
    End If
    Screen.MousePointer = vbDefault
    DizEnableEm True
    pgbProgress.Value = 1
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    DizEnableEm True
    MsgBox "Command3_Click in frmTest Error - " & Err.Number & ": " & Err.Description
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub DizEnableEm(Optional ByVal bEnabled As Boolean)
    cmdWriteU.Enabled = bEnabled
    cmdWriteS.Enabled = bEnabled
    cmdRandom.Enabled = bEnabled
    cmdRead.Enabled = bEnabled
    Command1.Enabled = bEnabled
    Command2.Enabled = bEnabled
    Command3.Enabled = bEnabled
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    cboState.ListIndex = 0
    GetFileUBs
    lMin = CLng(txtLB)
    lMax = CLng(cboUB.Text)
    cmdRandom_Click
    chkDirection_Click
    chkCaseSens_Click
    Set cSSort = New cStableSort
    Set cSSortCWP = New cStableSortCWP
    Set cSSorter = New cStableSorter
    cSSorter.Attach Me
    Exit Sub
ErrorHandler:
    MsgBox "Form_Load in frmTest Error - " & Err.Number & ": " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cSSort = Nothing
    Set cSSortCWP = Nothing
    cSSorter.Detach
    Set cSSorter = Nothing
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdSortClassInfo_Click()
    Dim ss As String
        ss = "The first sorting option utilizes the cStableSort " & _
           "class that is declared WithEvents, which raises a Compare " & _
           "event." & vbNewLine & vbNewLine & _
           "The following sorting option utilizes the cStableSorter " & _
           "class which Implements the ICompareClient interface." & vbNewLine & vbNewLine & _
           "Notice that the Implemented callback method is almost twice " & _
           "as fast as the Event raising method." & vbNewLine & vbNewLine & _
           "The last sorting option utilizes the cStableSortCWP class " & _
           "which uses a clever technique that " & _
           "I picked up from LukeH (selftaught) which uses the Win API " & _
           "CallWindowProc to fire a callback function to produce a " & _
           "Compare method in a standard module. This method is almost " & _
           "as fast as the Implements method." & vbNewLine & vbNewLine & _
           "Each solution provides a single 'Compare' routine within the " & _
           "client that you can code to compare any indexed data that is " & _
           "local to the client that references the sorting class." & vbNewLine & vbNewLine & _
           "All methods require an index (long) array to be passed " & _
           "instead of passing the source array. This gets around VB " & _
           "limitations passing some data types between classes." & vbNewLine & vbNewLine & _
           "Because they use an index array to sort the items no 'SwapItem' " & _
           "routine is required and so is considerably faster sorting the " & _
           "index array internally." & vbNewLine & vbNewLine & _
           "This permits smarter sorting instead of the limited item " & _
           "swap technique, which allows for fast copymemory operations " & _
           "that can manipulate multiple items simultaneously." & vbNewLine & vbNewLine
    Text1.Text = ss
    Me.Caption = " Generic Stable Indexed Data Sorting Classes"
End Sub

Private Sub cmdHelp_Click()
    Dim ss As String
        ss = "The easiest use is to build an array of random items " & _
           "and sort away." & vbNewLine & vbNewLine & _
           "You can save an unsorted 'built' array to file by " & _
           "clicking the 'write file (original)' button." & vbNewLine & vbNewLine & _
           "The UB combo will list all available test files that " & _
           "can be selected, then simply click 'read file'." & vbNewLine & vbNewLine & _
           "The idea to open an array is by setting the UB matching " & _
           "the file name e.g. 'Arr99999_U.dat' where the 99999 " & _
           "is the UB of a saved file in the ArrayFiles folder." & vbNewLine & vbNewLine & _
           "The _U in the file name signifies unsorted data. " & _
           "Saved files will be appended with _S if they are sorted." & vbNewLine & vbNewLine & _
           "You can import your own data by naming the file " & _
           "matching this format, with the number anything over " & _
           "the estimated size to fully load or it will truncate the " & _
           "file input at that UB. Blank lines in the file will be " & _
           "skipped so all array items will contain data." & vbNewLine & vbNewLine & _
           "The UB will be reset to the actual real UB of the array " & _
           "after loading the data." & vbNewLine & vbNewLine
    Text1.Text = ss
    Me.Caption = " Generic Stable Indexed Data Sorting Classes"
End Sub

Private Sub cmdResorterHlp_Click()
    Dim ss As String
        ss = "To appreciate the real benifit of the built-in Insert/Binary " & _
           "hybrid resorter you can do the following: " & vbNewLine & vbNewLine & _
           "First sort an array then click the 'write file (sorted)' " & _
           "button." & vbNewLine & vbNewLine & _
           "Notice the state combo will automatically be set to sorted." & vbNewLine & vbNewLine & _
           "You can then click the 'read file' button which will " & _
           "load the sorted file ready for re-sorting to compare " & _
           "refresh-sorting performance. " & vbNewLine & vbNewLine & _
           "In this same way you can reverse sort a sorted array, or " & _
           "turn case-sensitivity off, etc, to test the different " & _
           "sorting operations on pre-sorted data." & vbNewLine & vbNewLine & _
           "You can set the state combo at any time to " & _
           "read in a sorted or unsorted file." & vbNewLine & vbNewLine
    Text1.Text = ss
    Me.Caption = " Generic Stable Indexed Data Sorting Classes"
End Sub

Private Sub cmdUDTs_Click()
    Dim ss As String
        ss = "You can sort any data that is stored in an indexed data " & _
           "storage structure such as lists, collections, and arrays of " & _
           "all data types, including multi-dimensional arrays and arrays " & _
           "of UDTs." & vbNewLine & vbNewLine & _
           "It is up to you to write the comparison code needed in the " & _
           "exposed Compare routine to suit your particular storage structure " & _
           "and data type." & vbNewLine & vbNewLine & _
           "This is actually the best approach as it hides " & _
           "all the sorting details and leaves you to handle only the " & _
           "comparison code relavant to your current data implementation " & _
           "and desired sort criteria." & vbNewLine & vbNewLine & _
           "This is not intended to be a treatise on sorting multi-dimensional " & _
           "arrays or arrays of UDT's, if this is what you need then psc has " & _
           "has some very good examples on these subjects." & vbNewLine & vbNewLine & _
           "This is intended as a straighforward sorting solution that may " & _
           "be useful if you have a specific need to sort an indexed data structure, " & _
           "and you would like the sorting details to be hidden, and the sorting " & _
           "class's usage to be as simple and as generic as possible." & vbNewLine & vbNewLine & _
           "Each class header has clear instructions on how to implement " & _
           "that particular solution, and this form's code also demonstrates " & _
           "the usage of each sorting class without complicating it with data " & _
           "implementation details." & vbNewLine & vbNewLine
    Text1.Text = ss
    Me.Caption = " Generic Stable Indexed Data Sorting Classes"
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cboUB_LostFocus()
    If lMax < lMin Then txtLB = lMax
End Sub

Private Sub chkDirection_Click()
    eDirect = (chkDirection.Value * -2) + 1 '0 >> 1 : 1 >> -1
End Sub

Private Sub chkCaseSens_Click()
    ' vbBinaryCompare = 0, vbTextCompare = 1
    eCase = Abs(chkCaseSens.Value - 1) ' 1 >> 0 : 0 >> 1
    'eCase = Abs(Not CBool(chkCaseSens.Value))
    Call CWPMethod(eCase)
End Sub

Private Sub chkCaseSens_GotFocus()
    lPrevChkState = chkCaseSens.Value
End Sub

Private Sub chkCaseSens_LostFocus()
    If chkCaseSens.Value <> lPrevChkState Then ResetLabels
End Sub

Private Sub chkDirection_GotFocus()
    lPrevChkState = chkDirection.Value
End Sub

Private Sub chkDirection_LostFocus()
    If chkDirection.Value <> lPrevChkState Then ResetLabels
End Sub

Private Sub cboState_Click()
    suffix = "_" & Left$(cboState.Text, 1)
End Sub

Private Sub cmdClear_Click()
    ResetLabels
End Sub

Private Sub lblResults_Click(Index As Integer)
    ResetLabels Index
End Sub

Private Sub txtLB_Change()
    If IsNumeric(txtLB) Then
        lMin = CLng(txtLB)
    Else
        txtLB = 0
    End If
End Sub

Private Sub cboUB_Change()
    If IsNumeric(cboUB.Text) Then
        lMax = CLng(cboUB.Text)
    Else
        cboUB.Text = 0
    End If
End Sub

Private Sub cboUB_Click()
    lMax = CLng(cboUB.Text)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Text1.Top = 210
    Text1.Left = 3165
    pgbProgress.Top = Me.ScaleHeight - Text1.Top - 100
    Text1.Height = pgbProgress.Top - Text1.Top - 150
    pgbProgress.Width = Me.ScaleWidth - 300
    Text1.Width = Me.ScaleWidth - Text1.Left - 180
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdWriteU_Click()
    Dim FileNum As Integer, counter As Long
    Dim rc As eSortState
    On Error GoTo FileErrorHandler
    rc = strGetArraySortState(sA, lMin, lMax, eCase)
    If rc = Unsorted Then suffix = "_U" Else suffix = "_S"
    cboState.ListIndex = Abs(CBool(rc))
    Screen.MousePointer = vbHourglass
    FileNum = FreeFile
    Open App.Path & "\ArrayFiles\Arr" & lMax & suffix & ".dat" For Output As #FileNum
    For counter = lMin To lMax
        Print #FileNum, sA(counter)
    Next counter
    AddCombo lMax
FileErrorHandler:
    Close FileNum
    Screen.MousePointer = vbDefault
    If Err = 9 Then
        MsgBox "You haven't built or opened an array with this UB"
    ElseIf Err Then
        MsgBox "Error - " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub cmdWriteS_Click()
    Dim FileNum As Integer, counter As Long
    Dim rc As eSortState
    On Error GoTo ErrorHandler
    rc = strGetArraySortStateIndexed(sA, aIdx, lMin, lMax, eCase)
    If rc = Unsorted Then Err.Raise 9
    Screen.MousePointer = vbHourglass
    FileNum = FreeFile
    Open App.Path & "\ArrayFiles\Arr" & lMax & "_S.dat" For Output As #FileNum
    For counter = lMin To lMax
        Print #FileNum, sA(aIdx(counter))
    Next counter
    cboState.ListIndex = 1
    AddCombo lMax
ErrorHandler:
    Close FileNum
    Screen.MousePointer = vbDefault
    If Err = 9 Then
        MsgBox "You haven't sorted the current array since you " & vbNewLine & _
               "built it or opened it - no sorted data to save. "
    ElseIf Err Then
        MsgBox "Error - " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub cmdRead_Click()
    Dim FileNum As Integer, numcount As Long
    Dim s1 As String, sCapt As String
    Dim n1 As Long, n2 As Long
    Dim iOpt As Integer
    On Error GoTo ErrorHandler
    FileNum = FreeFile
    n2 = 99999999
    numcount = lMin
    ReDim sA(lMin To lMax) As String
    Screen.MousePointer = vbHourglass

    s1 = Dir$(App.Path & "\ArrayFiles\Arr" & lMax & suffix & ".dat")
    If s1 = vbNullString Then
        s1 = Dir$(App.Path & "\ArrayFiles\Arr" & lMax & "_?.dat")
        If s1 = vbNullString Then Err.Raise 53
    End If

    Open App.Path & "\ArrayFiles\" & s1 For Input As #FileNum
    Do While Not EOF(FileNum) And numcount <= lMax
        Line Input #FileNum, s1
        s1 = Trim$(s1)
        
        If s1 <> vbNullString Then
            If Len(s1) > n1 Then n1 = Len(s1)
            If Len(s1) < n2 Then n2 = Len(s1)

            sA(numcount) = s1
            numcount = numcount + 1
        End If
    Loop
    Close FileNum
    If lMax > numcount - 1 Then
        lMax = numcount - 1
        ReDim Preserve sA(lMin To lMax) As String
        cboUB.Text = lMax
    End If
    sCapt = " strs of varying len(" & n2 & " chars to " & n1 & " chars)"
    n1 = strGetArraySortState(sA, lMin, lMax, eCase)
    cboState.ListIndex = Abs(CBool(n1))
    If n1 = Unsorted Then
        sCapt = " unsorted" & sCapt
    ElseIf n1 = PreSorted Then
        sCapt = " pre sorted" & sCapt
    ElseIf n1 = PreRevSorted Then
        sCapt = " reverse sorted" & sCapt
    ElseIf n1 = MostlySorted Then
        sCapt = " mostly sorted" & sCapt
    ElseIf n1 = SemiSorted Then
        sCapt = " semi sorted" & sCapt
    ElseIf n1 = MostlyRevSorted Then
        sCapt = " mostly rev-sorted" & sCapt
    ElseIf n1 = SemiRevSorted Then
        sCapt = " semi rev-sorted" & sCapt
    End If
    Me.Caption = " LB=" & lMin & ", UB=" & lMax & sCapt

    AddCombo lMax
ErrorHandler:
    Close FileNum
    ReDim aIdx(lMin To lMax) As Long
    ResetIdxArray
    ResetCWPDemo sA
    ResetLabels
    cmdWriteS.Enabled = (n1 <> Unsorted)
    Screen.MousePointer = vbDefault
    If Err = 53 Then
        cmdHelp_Click
    ElseIf Err Then
        MsgBox "Error - " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub cmdRandom_Click()
    On Error GoTo ErrorHandler
    Me.Caption = " Sort Test"
    Dim n1 As Long, n2 As Long, j As Long
    n2 = 200000
    Screen.MousePointer = vbHourglass
    Randomize Timer
    ReDim sA(lMin To lMax) As String
    Dim s3(1 To 3) As String
    s3(1) = "Const AF_UNSPEC = 0         ' Unspecified. Although AF_UNSPEC is defined for backwards compatibility, using AF_UNSPEC for the af parameter when creating a socket is STRONGLY DISCOURAGED. The interpretation of the protocol parameter depends on the actual address family chosen. As environments grow to include more and more address families that use overlapping protocol values there is more and more chance of choosing an undesired address family when AF_UNSPEC is used."
    s3(2) = "Const EN_MAXTEXT = &H501    ' Notification message sent when the current text insertion has exceeded the specified number of characters for the edit control. The text insertion has been truncated. This message is also sent when an edit control does not have the ES_AUTOHSCROLL style and the number of characters to be inserted would exceed the width of the edit control. This message is also sent when an edit control does not have the ES_AUTOVSCROLL style and the total number of lines resulting from a text insertion would exceed the height of the edit control."
    s3(3) = "Declare Function CreateProcessAsUser Lib advapi32.dll Alias CreateProcessAsUserA (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long"
    For j = lMin To lMax
        sA(j) = Chr$(((j * Rnd) Mod 26) + 97) & " " & Chr$(((j * Rnd) Mod 26) + 97) & " " & Chr$(((j * Rnd) Mod 26) + 97) & "" & Chr$(((j * Rnd) Mod 26) + 97) & "" & Chr$(((j * Rnd) Mod 26) + 97) & " "
        sA(j) = sA(j) & Trim$(Mid$(s3(Abs(j) Mod 3 + 1), (Abs(j) * 44 * Rnd) Mod 470 + 1, (Abs(j) * 83 * Rnd) Mod 300 + 1))
        If j Mod 5 = 0 Then If j < lMax Then sA(j + 1) = UCase$(sA(j)): j = j + 1
        If Len(sA(j)) > n1 Then n1 = Len(sA(j))
        If Len(sA(j)) < n2 Then n2 = Len(sA(j))
    Next
    ReDim aIdx(lMin To lMax) As Long
    ResetIdxArray
    ResetCWPDemo sA
    ResetLabels
    cmdWriteS.Enabled = False
    cboState.ListIndex = 0
    Screen.MousePointer = vbDefault
    Me.Caption = " unsorted strs of varying len(" & n2 & " chars to " & n1 & " chars)"
    Me.Caption = " LB=" & lMin & ", UB=" & lMax & Me.Caption
    Exit Sub
ErrorHandler:
    MsgBox "cmdRandom_Click in frmTest Error - " & Err.Number & ": " & Err.Description
End Sub

Private Sub ResetIdxArray()
    Dim j As Long
    For j = lMin To lMax
        aIdx(j) = j
    Next
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++
