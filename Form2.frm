VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "��� ����"
   ClientHeight    =   6255
   ClientLeft      =   1710
   ClientTop       =   1785
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8280
   Begin VB.Frame Frame1 
      Caption         =   " ����� ���� "
      Height          =   840
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2325
      Width           =   7830
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "����� ������ ���� - ���� Ԙ�� - ���� ����� "
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   3705
      End
   End
   Begin MSComDlg.CommonDialog cdbColor 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMolahezeh 
      Caption         =   "&�������"
      Begin VB.Menu mnuSandoog 
         Caption         =   "�����"
         Begin VB.Menu mnuSodoor 
            Caption         =   "���� ���"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuAslah 
            Caption         =   "&����� ���"
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuPaziresh 
         Caption         =   "�����"
         Begin VB.Menu mnuTashkil 
            Caption         =   "����� ������ �����"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAslah_P 
            Caption         =   "����� ����� ������"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuKhoroog 
         Caption         =   "&����"
      End
   End
   Begin VB.Menu mnuCodes 
      Caption         =   "&����"
      Begin VB.Menu mnuCodeBimeh 
         Caption         =   "�� ����"
      End
      Begin VB.Menu mnuMasaref 
         Caption         =   "�� ����� ���"
      End
   End
   Begin VB.Menu mnuColor 
      Caption         =   "������� ��"
      Begin VB.Menu mnuTanzim 
         Caption         =   "����� �� �� �����"
      End
   End
   Begin VB.Menu mnuAmkanat 
      Caption         =   "&������� �����"
      Begin VB.Menu mnuMoShakh 
         Caption         =   "������ �����"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "����� ����"
      End
      Begin VB.Menu mnuPad 
         Caption         =   "���э� �������"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "�������"
         Begin VB.Menu mnutxtOpen 
            Caption         =   "��� ���� ��� �� ����� txt."
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public z As Variant

Private Sub Form_Activate()

     'cdbColor.CancelError = True
 
 'On Error GoTo dbErrHandler
   
  'cdbColor.Flags = cdlCCFullOpen + cdlCCHelpButton
'ColorDB
 ' cdbColor.ShowColor
  frmMain.BackColor = z
 ' Exit Sub
  
'dbErrHandler:
  
 ' Exit Sub
End Sub

Private Sub Form_Load()
   frmMain.Left = (Screen.Width - frmMain.Width) / 2
   frmMain.Top = (Screen.Height - frmMain.Height) / 2
   
    
   
   
   'Call Color
   
End Sub

Private Sub mnuAslah_Click()
   frmAslah.Show 1
End Sub

Private Sub mnuAslah_P_Click()
   frmAslah_P.Show 1
End Sub

Private Sub mnuCalc_Click()
   Shell ("calc.exe")
   
End Sub

Private Sub mnuCodeBimeh_Click()
   frmCodeBimeh.Show 1
End Sub

Private Sub mnuCodeDaroo_Click()
   frmDarooKhaneh.Show 1
End Sub

Private Sub mnuForoosh_Click()
   frmforoosh.Show 1
End Sub

Private Sub mnuKhoroog_Click()
   For intCtr = (Forms.Count - 1) To 0 Step -1
      Unload Forms(intCtr)  'Unloads both hidden and shown forms
   Next intCtr
End Sub

Private Sub mnuMasaref_Click()
   frmCodeMasaref.Show 1
End Sub


Private Sub mnuMoShakh_Click()
   frmAbout.Show 1
   
End Sub

Private Sub mnuPad_Click()
   Shell ("NOTEPAD.exe")
End Sub

Private Sub mnuSodoor_Click()
   frmSodoor.Show 1
End Sub

Private Sub form1_Click()
   Form3.Show 1
End Sub

Private Sub mnuTanzim_Click()
   cdbColor.CancelError = True
   
   On Error GoTo dbErrHandler
   
  'cdbColor.Flags = cdlCCFullOpen + cdlCCHelpButton
  
 
  cdbColor.Flags = 2 + 8 + 1
'ColorDB
  cdbColor.ShowColor
  z = cdbColor.Color
  frmMain.BackColor = z
  Exit Sub
  
dbErrHandler:
  
  Exit Sub
End Sub

Private Sub mnuTashkil_Click()
   frmTashkilParv.Show 1
End Sub

Private Sub mnutxtOpen_Click()
   cdbColor.FileName = " "
   cdbColor.Filter = "*.txt|*.txt"
   cdbColor.ShowOpen
   Shell "notepad " & cdbColor.FileName
End Sub
