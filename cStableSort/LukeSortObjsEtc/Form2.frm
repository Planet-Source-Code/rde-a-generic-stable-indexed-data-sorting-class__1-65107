VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ShellSortObject"
   ClientHeight    =   4770
   ClientLeft      =   3465
   ClientTop       =   795
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7650
   Begin VB.CommandButton Command2 
      Caption         =   "Sort Longs"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtItems 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Text            =   "1000"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ReDim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1200
      TabIndex        =   3
      Top             =   435
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form2.frx":0000
      Left            =   2760
      List            =   "Form2.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4740
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   96
      TabIndex        =   0
      Top             =   1056
      Width           =   7404
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "Sort Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim Employees() As CEmployee

Private Sub Combo1_Click()
    Dim startTime As Single
    Dim sortMethod As Integer
    
    sortMethod = Combo1.ListIndex
    If sortMethod = 0 Then Exit Sub
    
    startTime = Timer
    Screen.MousePointer = vbHourglass
    lblTime = ""
    
    If sortMethod = 1 Then
        ShellSortAny VarPtr(Employees(1)), UBound(Employees), 4&, AddressOf CompareName
    ElseIf sortMethod = 2 Then
        ShellSortAny VarPtr(Employees(1)), UBound(Employees), 4&, AddressOf CompareDeptName
    ElseIf sortMethod = 3 Then
        ShellSortAny VarPtr(Employees(1)), UBound(Employees), 4&, AddressOf CompareSalaryName
    ElseIf sortMethod = 4 Then
        ShellSortAny VarPtr(Employees(1)), UBound(Employees), 4&, AddressOf CompareDeptSalaryName
    End If
    
    Screen.MousePointer = vbDefault
    If Me.Visible Then
        lblTime = Format$(Timer - startTime, "#0.00") & " secs."
    End If
    RefreshList
End Sub

Private Sub Command2_Click()
    Dim liCount As Long
    Dim liLongs() As Long
    liCount = Val(txtItems.Text)
    Randomize Timer
    ReDim liLongs(0 To liCount)
        
    For liCount = 0 To liCount
        liLongs(liCount) = Rnd * 428000000 - 214000000
    Next
    
    If MsgBox("Random Array Populated.  Press OK to sort, or cancel to cancel.", vbOKCancel) = vbCancel Then Exit Sub
    
    ShellSortAny VarPtr(liLongs(0)), liCount, 4&, AddressOf CompareLong
    
    If MsgBox("Array Sorted!  Display results in the immediate window?", vbYesNo) = vbYes Then
        For liCount = 0 To liCount - 1&
            Debug.Print liLongs(liCount)
        Next
    End If
    
End Sub

Private Sub Form_Initialize()
    ' initialize the array
    ReDim Employees(1 To 1) As CEmployee
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    ' set tabstops in List1
    Dim tabs(3) As Long
    tabs(1) = 80
    tabs(2) = 160
    SendMessage List1.hWnd, LB_SETTABSTOPS, 2, tabs(1)
End Sub


Private Sub Command1_Click()
    ' add 100 random items
    Dim i As Integer
    Dim numEls As Long
    
    numEls = Val(txtItems)
    ReDim Preserve Employees(1 To numEls) As CEmployee
    
    For i = 1 To numEls
        Set Employees(i) = New CEmployee
        With Employees(i)
            .FirstName = Choose(Rnd * 10 + 1, "Joe", "Robert", "Frank", "Anne", "Jim", "Nicole", "Michael", "George", "Nina", "John")
            .LastName = Choose(Rnd * 10 + 1, "Smith", "Ford", "Aiello", "Green", "Goldson", "Ryan", "Tracey", "Simone", "Halen", "Douglas")
            .Dept = Choose(Rnd * 5 + 1, "Sales", "Management", "Marketing", "Cust.Service", "Purchases")
            .Salary = Int(Rnd * 80 + 20) * 1000
        End With
    Next
    
    Combo1.ListIndex = 0
    RefreshList
    
End Sub

Private Sub Form_Resize()
    Combo1.Width = ScaleWidth - Combo1.Left
    List1.Width = ScaleWidth - List1.Left * 2
    List1.Height = ScaleHeight - List1.Top
End Sub

Private Sub RefreshList()
    Dim index As Long
    
    List1.Clear
    For index = 1 To UBound(Employees)
        With Employees(index)
            List1.AddItem .ReversedName & vbTab & .Dept & vbTab & .Salary
        End With
    Next
End Sub

