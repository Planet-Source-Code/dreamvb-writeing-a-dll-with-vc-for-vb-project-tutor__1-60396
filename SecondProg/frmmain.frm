VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "VC Dll Test"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "LCaseEx"
      Height          =   465
      Left            =   3180
      TabIndex        =   7
      Top             =   255
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   465
      Left            =   3180
      TabIndex        =   6
      Top             =   840
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SHR 16,2"
      Height          =   465
      Left            =   1695
      TabIndex        =   5
      Top             =   1425
      Width           =   1380
   End
   Begin VB.CommandButton cmdshl 
      Caption         =   "SHL 8,2"
      Height          =   465
      Left            =   1695
      TabIndex        =   4
      Top             =   840
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pass Array"
      Height          =   465
      Left            =   1695
      TabIndex        =   3
      Top             =   210
      Width           =   1380
   End
   Begin VB.CommandButton cmducase 
      Caption         =   "UcaseEx"
      Height          =   465
      Left            =   105
      TabIndex        =   2
      Top             =   1410
      Width           =   1380
   End
   Begin VB.CommandButton cmdRepVer 
      Caption         =   "Report Version"
      Height          =   465
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   1410
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add Two Numbers"
      Height          =   465
      Left            =   105
      TabIndex        =   0
      Top             =   765
      Width           =   1410
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReportVersion Lib "vbTest.dll" () As Long
Private Declare Function Add Lib "vbTest.dll" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub UcaseEx Lib "vbTest.dll" (ByVal s As String)
Private Declare Function LCaseEx Lib "vbTest.dll" (ByVal s As String) As String

Private Declare Function Sum Lib "vbTest.dll" (ByRef slots As Long, ByVal nSize As Long) As Long
Private Declare Function SHL Lib "vbTest.dll" (ByVal x As Long, ByVal places As Long) As Long
Private Declare Function SHR Lib "vbTest.dll" (ByVal x As Long, ByVal places As Long) As Long

' Hi this is a simple project of how to work with DLL Files from Visual C++
' example includes how to pass Numeric and strings and return results.
' also shows how to pass an array
' Once you compile the Visual C++ dll place it into the same folder as this project

'Note That:
' SHL and SHR functions don;t seem to work when ran in the VB IDE so it's best to compile
' the project then run it. not sure why this happens eveything else works fine

Private Sub cmdadd_Click()
Dim x As Integer, y As Integer
    ' example to add two numbers
    x = 10
    y = 15
    MsgBox "Result of " & x & "+" & y & "=" & Add(x, y)
    
End Sub

Private Sub cmdshl_Click()
    ' Bitshift left 8 by 2 places
    MsgBox "SHL 8,2 " & vbCrLf & "Result: " & SHL(8, 2)
End Sub

Private Sub Command1_Click()
Dim TestArray(4) As Long
Dim nResult As Long
    'OK what this little example does is call the Sum Function in vbTest.dll
    ' To sum up all of the TestArray elements
    TestArray(0) = 8
    TestArray(1) = 16
    TestArray(2) = 32
    TestArray(3) = 64
    TestArray(4) = 128
    
    nResult = Sum(TestArray(0), 5) ' result should be 248
    MsgBox "Result is : " & nResult
End Sub

Private Sub cmdRepVer_Click()
    Dim nResult As Long
    ' This example returns the version of the dll we set
    ' if all goes well result will be 1
    nResult = ReportVersion
    MsgBox "ReportVersion : " & nResult
End Sub

Private Sub Command2_Click()
    ' Bitshift right 16 by 2 places
    MsgBox "SHR 16,2 " & vbCrLf & "Result: " & SHR(16, 2)
End Sub

Private Sub cmducase_Click()
Dim m_str As String
    ' Example of passing a string and returning it's result
    ' for an example I written an functionto upper case the text passed.
    ' Just top give you an idea
    m_str = InputBox("Enter some text to uppercase:")
    If Len(Trim(m_str)) = 0 Then m_str = ""
    Call UcaseEx(m_str)
    MsgBox "Result: " & m_str
    
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
Dim s As String, b As String
    'Lower case a string
    s = "HELLO WORLD"
    
    b = "Orginal String: " & s & vbCrLf
    s = b & "UcaseEx Result: " & LCase(s)
    MsgBox s, vbInformation
    
    b = ""
    s = ""
    
End Sub
