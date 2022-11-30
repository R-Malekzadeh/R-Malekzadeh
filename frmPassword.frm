VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    'check for correct password
    'If txtPassword = "password" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
    '    LoginSucceeded = True
    '    Me.Hide
    'Else
    '    MsgBox "Invalid Password, try again!", , "Login"
    '    txtPassword.SetFocus
    '    SendKeys "{Home}+{End}"
    'End If
   Dim p As Integer
   p = p + 1
If p > 3 Then
   MsgBox "Sorry,You coulden't continue!", , "Notice"
   End
ElseIf txtPassword.Text = "111" Then
       frmPassword.Hide
       frmMain.Show vbModal, Me
       End
    Else
       MsgBox "Invalid Password, try again!", , "Notice"
       txtPassword.Text = ""
       txtPassword.SetFocus
End If



End Sub

Private Sub Form_Load()
   frmPassword.Left = (Screen.Width - frmPassword.Width) / 2
   frmPassword.Top = (Screen.Height - frmPassword.Height) / 2
   cmdOK.Enabled = False
   txtUserName.Text = "PROJECT"
End Sub

Private Sub txtPassword_Change()
   If txtPassword.Text <> "" Then
   cmdOK.Enabled = True
Else
   cmdOK.Enabled = False
End If
End Sub

Private Sub txtUserName_Change()
   txtUserName.Enabled = False
End Sub
