VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VC Test"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   1545
      TabIndex        =   7
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   3165
      TabIndex        =   3
      Top             =   465
      Width           =   1215
   End
   Begin VB.TextBox txtB 
      Height          =   495
      Left            =   1665
      TabIndex        =   2
      Top             =   465
      Width           =   1215
   End
   Begin VB.TextBox txtA 
      Height          =   495
      Left            =   195
      TabIndex        =   1
      Top             =   465
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   345
      Left            =   195
      TabIndex        =   0
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Answer"
      Height          =   210
      Left            =   3240
      TabIndex        =   6
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Input2"
      Height          =   210
      Left            =   1680
      TabIndex        =   5
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Input1"
      Height          =   210
      Left            =   255
      TabIndex        =   4
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Add Lib "VBTest.dll" (ByVal x As Long, ByVal y As Long) As Long

Private Sub cmdAdd_Click()
Dim A As Long, B As Long
    'Add two numbers
    A = 0: B = 0: Result = 0
    
    A = Val(txtA.Text)
    B = Val(txtB.Text)
    
    txtResult.Text = Add(A, B)
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub
