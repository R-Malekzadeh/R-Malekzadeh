VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTashkilParv 
   Caption         =   "›—„  ‘ﬂÌ· Å—Ê‰œÂ"
   ClientHeight    =   6555
   ClientLeft      =   1395
   ClientTop       =   1440
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   8760
   Begin MSAdodcLib.Adodc adoTashkilParv 
      Height          =   330
      Left            =   120
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\main rational\Project_vb111\DB1.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\main rational\Project_vb111\DB1.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblTashkil"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "’œÊ—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   7935
      Begin VB.TextBox txtJensyat 
         Alignment       =   1  'Right Justify
         DataField       =   "Jensyat"
         DataSource      =   "adoTashkilParv"
         Height          =   285
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtComment 
         DataField       =   "Comment"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   240
         MaxLength       =   300
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2160
         Width           =   7575
      End
      Begin VB.TextBox txtAllat 
         DataField       =   "Alat"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtLName 
         DataField       =   "Last_Name"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtFName 
         DataField       =   "Frist_Name"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtSenn 
         DataField       =   "Senn"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   5040
         TabIndex        =   24
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton optMoanas 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ê‰À"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optMozakar 
            Alignment       =   1  'Right Justify
            Caption         =   "„–ﬂ—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄·  „—«Ã⁄Â"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê÷ÌÕ« :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "”‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã‰”"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   7935
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "13##/##/##"
         PromptChar      =   " "
      End
      Begin VB.ComboBox cmbMorajeh 
         DataField       =   "Morajeh"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTashkilParv.frx":0000
         Left            =   600
         List            =   "frmTashkilParv.frx":000A
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtShomP 
         DataField       =   "Shom_Parvandeh"
         DataSource      =   "adoTashkil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "¬Ì« ﬁ»·« „—«Ã⁄Â ﬂ—œÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ „—«Ã⁄Â"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   6180
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   2
            Picture         =   "frmTashkilParv.frx":0018
            TextSave        =   "03:50 ⁄’—"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "03/01/1383"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmTashkilParv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con2 As New ADODB.Connection
Private Sub chkMoanas_LostFocus()
'    If chkMoanas.Value = 1 Then
  ' Check1.Enabled = False
' Check2.Value = 1
'  ChkMozakar.Value = 0
'Else
' If ChkMozakar.Value = 1 Then
'  chkMoanas.Value = 0
'End If
'End If
If (chkMoanas.Value = 1) And (ChkMozakar.Value = 1) Then
  ' Check1.Enabled
 MsgBox ("‰„Ì  Ê«‰Ìœ Â— œÊ ê“Ì‰Â —« «‰ Œ«» ò‰Ìœ"), vbCritical
 
  chkMoanas.Value = 0
  ChkMozakar.Value = 0

 ChkMozakar.SetFocus
End If

End Sub

Private Sub cmbMorajeh_GotFocus()
   
  'txtTarikh.Text = CDate(txtTarikh.Text)
End Sub

Private Sub cmdSave_Click()
  Dim rs As New ADODB.Recordset


    
  If CheckDate(MaskEdBox1) = False Then
    MsgBox " «—ÌŒ «‘ »«Â «” "
    MaskEdBox1.SetFocus
    Exit Sub
   End If
   


If Len(cmbMorajeh) = 0 Then
    MsgBox "ò«œ— „—«Ã⁄Â —« Ê«—œ ò‰Ìœ"
    cmbMorajeh.SetFocus
    Exit Sub
End If

If (optMozakar.Value = False) And (optMoanas.Value = False) Then
    MsgBox " Ã‰”Ì  —« Ê«—œ ò‰Ìœ"
    optMozakar.SetFocus
    Exit Sub
End If
'
If Len(txtSenn) = 0 Then
    MsgBox "”‰ —« Ê«—œ ò‰Ìœ"
    txtSenn.SetFocus
    Exit Sub
End If


If Len(txtFName) = 0 Then
    MsgBox "‰«„ —« Ê«—œ ò‰Ìœ"
    txtFName.SetFocus
    Exit Sub
End If

If Len(txtLName) = 0 Then
    MsgBox "‰«„ Œ«‰Ê«œêÌ —« Ê«—œ ò‰Ìœ"
    txtLName.SetFocus
    Exit Sub
End If

If Len(txtAllat) = 0 Then
    MsgBox "⁄·  „—«Ã⁄Â —« Ê«—œ ò‰Ìœ"
    txtAllat.SetFocus
    Exit Sub
End If
'Sql = "Select * From tblkindbimeh where name='" & DcmbKindBimeh & "'"
'rs.Open Sql, con
'codeb = rs!code_Bimeh
'rs.Close

rs.Open "select *from tblTashkil", con2, adOpenKeyset, adLockOptimistic
rs.AddNew
rs!Shom_Parvandeh = Val(Label1)
rs!Tarikh_Morajeh = MaskEdBox1
rs!Morajeh = cmbMorajeh
'rs!Mozakar = optMozakar
'rs!Moanas = optMoanas
rs!Jensyat = txtJensyat
rs!Senn = Val(txtSenn)
rs!Frist_Name = txtFName
rs!Last_Name = txtLName
rs!Allat = txtAllat
rs!Comment = txtComment
rs.Update

rs.Close
txtShomP.Enabled = False
Label1 = Val(Label1) + 1
txtShomP = Label1
'MaskEdBox1 = "13  / /  "
cmbMorajeh = ""
optMoanas.Value = False
optMozakar.Value = False
txtJensyat = ""
txtSenn = ""
txtFName = ""
txtLName = ""
txtAllat = ""
txtComment = ""
MaskEdBox1.SetFocus

End Sub

Private Sub Form_Load()
   frmTashkilParv.Left = (Screen.Width - frmTashkilParv.Width) / 2
   frmTashkilParv.Top = (Screen.Height - frmTashkilParv.Height) / 2
   
   Dim h As String
   
   h = App.Path
   adoTashkilParv.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & h & "\DB1.MDB"
   adoTashkilParv.RecordSource = "tblTashkil"
   adoTashkilParv.Refresh

   con2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & h & "\DB1.MDB"
   
   Dim rs As New ADODB.Recordset
   
  rs.Open "select *from tblTashkil", con2
  If rs.BOF = True And rs.EOF = True Then
     
     MsgBox "ÃœÊ· ﬁ»Ê÷ Œ«·Ì «”  ", vbSystemModal
        
  End If
  rs.Close
  Dim sq As String
'  rs.Open "Select max(Shom_Parvandeh) as mShom_Parvandeh from tblTashkil ", con2
'  If rs.BOF = False And rs.EOF = False Then
'     If IsNull(rs!mShom_Parvandeh) Then
'         Label1 = 1
'     Else
'        Label1 = rs!mShom_Parvandeh + 1
'     End If
'  End If
'  rs.Close

sq = "Select max(Shom_Parvandeh) as mShom_Parvandeh from tblTashkil" '  where Tarikh_Morajeh= ' " & "'"
'& MaskEdBox1 & "'"
rs.Open sq, con2
  If rs.BOF = False And rs.EOF = False Then
     If IsNull(rs!mShom_Parvandeh) Then
         Label1 = 1
         txtShomP = Label1
         txtShomP.Enabled = False
     Else
        Label1 = rs!mShom_Parvandeh + 1
        txtShomP.Enabled = False
        txtShomP = Label1
     End If
  End If
  rs.Close
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   con2.Close
End Sub

Private Sub txtTarikh_Change()
   'txtTarikh.Text = Format(Now, "Date")
   'a = Date
   'txtTarikh.Text = Date
   'txtShomP.Text = CDate(txtShomP.Text)
   'txtTarikh.Text = CDate(txtShomP.Text)
   'If Not(txtTarikh.Text = "") Then
   '   txtTarikh.Text = CDate(txtTarikh.Text)
   'End If
   'txtTarikh.Text =
End Sub

Private Sub Label1_Click()

'Dim rs  As New ADODB.Recordset
'Sq = "Select * From tblTashkil where Tarikh_Morajeh'" & MaskEdBox1 & "'"
'rs.Open Sq, con2
'a = rs!Tarikh_Morajeh
'If MaskEdBox1 = a Then
'   i = i + 1
'End If
'
'Exit Sub
'rs.Close
End Sub

Private Sub MaskEdBox1_LostFocus()
   If CheckDate(MaskEdBox1) = False Then
    MsgBox " «—ÌŒ «‘ »«Â «” "
    MaskEdBox1.SetFocus
   End If
End Sub

Private Sub optMoanas_Click()
  ' Dim rs As New ADODB.Recordset
 '  rs.Open "select Moanas from tblTashkil"
   
   If optMoanas.Value = True Then
 '     rs!Moanas = optMoanas
      txtJensyat.Text = "„Ê‰À"
   Else
      txtJensyat.Text = ""
   End If
  'txtSenn.SetFocus
 '  rs.Close
   
End Sub

Private Sub optMozakar_Click()
   If optMozakar.Value = True Then
    '  rs!Mozakar = optMozakar
      txtJensyat.Text = "„–ﬂ—"
   Else
      txtJensyat.Text = ""
   
   End If
End Sub
