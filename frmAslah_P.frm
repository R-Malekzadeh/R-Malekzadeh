VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAslah_P 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAslah_P.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Shom_Parvandeh"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Tarikh_Morajeh"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Morajeh"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Jensyat"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Senn"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Frist_Name"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Last_Name"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Allat"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Comment"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Õ–›"
      Height          =   375
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Ã” ÃÊ"
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adoAslah_P 
      Height          =   330
      Left            =   0
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAslah 
      Caption         =   "«’·«Õ"
      Height          =   375
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   7935
      Begin VB.TextBox txtTarikh 
         Alignment       =   1  'Right Justify
         DataField       =   "Tarikh_Morajeh"
         DataSource      =   "adoAslah_P"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtShomP 
         DataField       =   "Shom_Parvandeh"
         DataSource      =   "adoAslah_P"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbMorajeh 
         DataField       =   "Morajeh"
         DataSource      =   "adoAslah_P"
         Height          =   315
         ItemData        =   "frmAslah_P.frx":0019
         Left            =   600
         List            =   "frmAslah_P.frx":0023
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   720
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Å—Ê‰œÂ"
         Height          =   255
         Left            =   6720
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ „—«Ã⁄Â"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "¬Ì« ﬁ»·« „—«Ã⁄Â ﬂ—œÂ"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3480
      Width           =   7935
      Begin VB.TextBox txtJensyat 
         Alignment       =   1  'Right Justify
         DataField       =   "Jensyat"
         DataSource      =   "adoAslah_P"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         TabIndex        =   11
         Top             =   120
         Width           =   1455
         Begin VB.OptionButton optMozakar 
            Alignment       =   1  'Right Justify
            Caption         =   "„–ﬂ—"
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optMoanas 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ê‰À"
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.TextBox txtSenn 
         DataField       =   "Senn"
         DataSource      =   "adoAslah_P"
         Height          =   285
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFName 
         DataField       =   "Frist_Name"
         DataSource      =   "adoAslah_P"
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtLName 
         DataField       =   "Last_Name"
         DataSource      =   "adoAslah_P"
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtAllat 
         DataField       =   "Allat"
         DataSource      =   "adoAslah_P"
         Height          =   315
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtComment 
         DataField       =   "Comment"
         DataSource      =   "adoAslah_P"
         Height          =   675
         Left            =   240
         MaxLength       =   300
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2040
         Width           =   7575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã‰”"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "”‰"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê÷ÌÕ« :"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„"
         Height          =   255
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄·  „—«Ã⁄Â"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAslah_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Private Sub cmbMorajeh_Change()
   If cmbMorajeh.Tag <> cmbMorajeh.Text Then
      Call TextChanged
   End If
End Sub

Private Sub cmdAslah_Click()
   Dim a As String
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
   a = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «’·«Õ ﬂ‰Ìœ", vbYesNo)
   If a = vbYes Then
      rs.Open "select *from tblTashkil", con1
      adoAslah_P.Recordset.Update
      rs.Close
      
   End If
   
      
End Sub

Private Sub cmdDel_Click()
   Dim a As String
   Dim rs As New ADODB.Recordset
   rs.Open "select *from tblTashkil", con1
   
   If rs.BOF = True And rs.EOF = True Then
      MsgBox "ÃœÊ· Œ«·Ì «” ", vbOKOnly
   Else
     ' If rs.BOF = True Or rs.EOF = True Then
         
         a = MsgBox("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ «Ì‰ —ﬂÊ—œ —« Õ–› ﬂ‰Ìœø", vbQuestion + vbYesNo, _
                 "Õ–› —ﬂÊ—œ")
         If a = vbYes Then
          adoAslah_P.Recordset.Delete
          
         End If
     ' End If
      
   End If
   rs.Close
End Sub

Private Sub cmdSearch_Click()
   
   If txtFName = "" Then
    adoAslah_P.RecordSource = "select *from tblTashkil "
    adoAslah_P.Refresh
    DataGrid1.ReBind
Else

    adoAslah_P.RecordSource = "select *from tblTashkil where tblTashkil.Frist_Name = '" & txtFName & " ' "
    '"select tblTashkil.Shom_Parvandeh, tblTashkil.Tarikh_Morajeh, tblTashkil.Morajeh, tblTashkil.Mozakar, tblTashkil.Moanas,tblTashkil.Senn,tblTashkil.Frist_Name,tblTashkil.Last_Name,tblTashkil.Allat,tblTashkil.Comment where tblTashkil.Frist_Name = '" & txtFName & " ' "
    adoAslah_P.Refresh
    DataGrid1.ReBind
End If
End Sub


Private Sub Command1_Click()
      
End Sub

Private Sub DataGrid1_Click()
  ' Dim rs As New ADODB.Recordset
   
  ' rs.Open "select * from tblTashkil", con1
   
  ' If rs!Moanas = True Then
  '    optMoanas.Value = True And optMozakar.Value = False
   
     
   
  ' ElseIf rs!Moanas = False Then
  '    optMoanas.Value = False And optMozakar.Value = True
   
  ' End If
   
  ' rs.Close
End Sub

Private Sub Form_Load()
   
   frmAslah_P.Left = (Screen.Width - frmAslah_P.Width) / 2
   frmAslah_P.Top = (Screen.Height - frmAslah_P.Height) / 2

   adoAslah_P.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\DB1.MDB"
   adoAslah_P.RecordSource = "select tblTashkil.Shom_Parvandeh, tblTashkil.Tarikh_Morajeh, tblTashkil.Morajeh,tblTashkil.Jensyat,tblTashkil.Senn,tblTashkil.Frist_Name, tblTashkil.Last_Name, tblTashkil.Allat, tblTashkil.Comment from tblTashkil"
   adoAslah_P.Refresh
  
   con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db1.mdb"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   con1.Close
End Sub
Private Sub TextChanged()
   'If adoTel.Recordset.EOF Or adoTel.Recordset.BOF Then
      ' Do Nothing
   'End If
End Sub



Private Sub MaskEdBox1_LostFocus()
   b = MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ  «—ÌŒ —«  €ÌÌ— œÂÌœø", vbYesNo)
   
   If b = vbYes Then
   
   If CheckDate(MaskEdBox1) = False Then
    MsgBox " «—ÌŒ «‘ »«Â «” "
    MaskEdBox1.SetFocus
    Exit Sub
   End If
   a = MaskEdBox1
   txtTarikh.Text = a
  'txtTarikh.Visible = True
   txtTarikh.Visible = True
   txtTarikh.Enabled = True
   txtTarikh.SetFocus
   cmbMorajeh.SetFocus
   txtTarikh.Enabled = False
   
   txtTarikh.Visible = False
   End If
End Sub

Private Sub optMoanas_Click()
   If optMoanas.Value = True Then
 '     rs!Moanas = optMoanas
      txtJensyat.Visible = True
      txtJensyat.Enabled = True
      txtJensyat.Text = "„Ê‰À"
      txtJensyat.SetFocus
      txtJensyat.Enabled = False
      txtJensyat.Visible = False
   Else
      txtJensyat.Text = ""
   End If
End Sub

Private Sub optMozakar_Click()
   If optMozakar.Value = True Then
    '  rs!Mozakar = optMozakar
      txtJensyat.Visible = True
      txtJensyat.Enabled = True
      txtJensyat.Text = "„–ﬂ—"
      txtJensyat.SetFocus
      txtJensyat.Enabled = False
      txtJensyat.Visible = False
   Else
      txtJensyat.Text = ""
   
   End If
End Sub

Private Sub txtAllat_Change()
   If txtAllat.Tag <> txtAllat.Text Then
      Call TextChanged
   End If
End Sub

Private Sub txtComment_Change()
   If txtComment.Tag <> txtComment.Text Then
      Call TextChanged
   End If
End Sub

Private Sub txtFName_Change()
   If txtFName.Tag <> txtFName.Text Then
      txtFName.Tag = txtFName.Text
   End If
End Sub

Private Sub txtLName_Change()
   If txtLName.Tag <> txtLName.Text Then
      Call TextChanged
   End If
End Sub

Private Sub txtSenn_Change()
'   If txtSenn.Tag <> txtSenn.Text Then
'      Call TextChanged
'   End If
End Sub



Private Sub txtShomP_Change()
  ' If txtShomP.Tag <> txtShomP.Text Then
  '    Call TextChanged
  ' End If
End Sub

'Private Sub DataReposition()
'  Dim rs As New ADODB.Recordset
'  rs.Open "select *from tblTashkil", con1
'
'   With adoAslah_P.Recordset.MoveNext
'      If .EOF Then
'         .MovePrevious
'      End If
'   End With
'   rs.Close
'End Sub
Private Sub txtTarikh_Change()
   txtTarikh.Text = MaskEdBox1
End Sub
