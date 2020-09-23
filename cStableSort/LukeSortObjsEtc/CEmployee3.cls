VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------
'    CEmployee Class
'
' This is a simple class, whose instances are meant
' to be collected in the CEmployees collection class
'-------------------------------------------------------------

Option Explicit

Public ID As Long
Public FirstName As String
Public LastName As String
Public Dept As String
Public Salary As Currency
Public Hired As Date

Public Property Get Name() As String
    Name = FirstName & " " & LastName
End Property

Public Property Get ReversedName() As String
Attribute ReversedName.VB_UserMemId = 0
    ReversedName = LastName & ", " & FirstName
End Property

