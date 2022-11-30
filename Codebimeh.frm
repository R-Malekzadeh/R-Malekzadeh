VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmCodeBimeh 
   Caption         =   "›—„ ﬂœ »Ì„Â"
   ClientHeight    =   6630
   ClientLeft      =   555
   ClientTop       =   1350
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   7695
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   7695
         Begin VB.TextBox txtCodeBimeh 
            Alignment       =   1  'Right Justify
            DataField       =   "code_Bimeh"
            DataSource      =   "adoCodeBimeh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  'Right Justify
            DataField       =   "Name"
            DataSource      =   "adoCodeBimeh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4680
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtDarsad 
            Alignment       =   1  'Right Justify
            DataField       =   "Darsad"
            DataSource      =   "adoCodeBimeh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   915
            MaxLength       =   3
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   390
            Width           =   2250
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂœ »Ì„Â"
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
            Index           =   0
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "‰«„ »Ì„Â"
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
            TabIndex        =   12
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "œ—’œ »Ì„Â"
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
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Label lblCodeBimeh 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ »Ì„Â"
         Height          =   135
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Õ–› ﬂ—œ‰"
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
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Codebimeh.frx":0000
      Height          =   2325
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4101
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "·Ì”  »Ì„Â Â«"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "code_Bimeh"
         Caption         =   "ﬂœ »Ì„Â"
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
         DataField       =   "Name"
         Caption         =   "‰«„ »Ì„Â"
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
         DataField       =   "Darsad"
         Caption         =   "œ—’œ"
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCodeBimeh 
      Height          =   330
      Left            =   120
      Top             =   6120
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.CommandButton cmdExit1 
      Caption         =   "&Œ—ÊÃ"
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
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "À»  ﬂ—œ‰"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "«÷«›Â ﬂ—œ‰"
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
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   ":‰«„ Ê ‰«„ Œ«‰Ê«œêÌ „”∆Ê· ¬„«œÂ ”«“Ì"
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
      Index           =   1
      Left            =   5040
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmCodeBimeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim con As New ADODB.Connection

Private Sub cmdDelete_Click()
DataGrid1.EditActive = True
txtCodeBimeh.Enabled = True
txtName.Enabled = True
txtDarsad.Enabled = True
If ((txtCodeBimeh = " ") And (txtName = " ") And (txtDarsad = " ")) Then
   MsgBox "”ÿ— „Ê—œ‰Ÿ— —« «‰ Œ«» ﬂ‰Ìœ"
End If
End Sub

Private Sub cmdExit1_Click()
  'On Error GoTo dbErrHandler
  'Dim rs As New ADODB.Recordset
  'rs.Open "Select * From tblKindBimeh", con, adOpenKeyset, adLockOptimistic
  
   
  If cmdExit1.Caption = "&Œ—ÊÃ" Then
      Unload Me
     ' GoTo dbErrHandler
  End If
   'Else
      
   '   rs.CancelUpdate
     ' If rs.EditMode <> adEditAdd Then
     '    rs.Move (0)
      'GoTo dbErrHandler
      'Else
    
     '   txtCodeBimeh.Enabled = " "
     '    txtName.Enabled = " "
     '    txtDarsad.Enabled = " "
        ' GoTo dbErrHandler
  ' End If
    'Else
    '   MsgBox "œ— Õ«· «ÌÃ«œ —ﬂÊ—œ «”  ", vbInformation
    '  End If
  '  txtCodeBimeh.Enabled = False
  '  txtName.Enabled = False
  '  txtDarsad.Enabled = False
   'GoTo dbErrHandler
   'End If
   
   'Call ToggleButtons
'dbErrHandler:
End Sub

Private Sub cmdNew_Click()
   adoCodeBimeh.Recordset.AddNew
   txtCodeBimeh.Enabled = True
   txtName.Enabled = True
   txtDarsad.Enabled = True
   txtCodeBimeh.SetFocus
'   Call ToggleButtons
End Sub

Private Sub cmdSave_Click()
  ' Dim boolAdding As Boolean
  'Dim mboolError As Boolean
   'mboolError = False
   'boolAdding = (adoSodoor.Recordset.EditMode = adEditAdd)

   'adoSodoor.Recordset.Update
   'If Not mboolError Then
     ' If boolAdding Then
     '    adoSodoor.Recordset.MoveLast
     ' End If
     ' Call ToggleButtons
     ' cmdSave.Enabled = False
  ' End If
'  Dim rs As New ADODB.Recordset
  
'  rs.Open "Select * From tblKindBimeh", con, adOpenKeyset, adLockOptimistic
'  rs!code_Bimeh = txtCodeBimeh
'  rs!Name = txtName
'  rs!Darsad = txtDarsad
'  rs.Update
'  rs.Close
  txtCodeBimeh.Enabled = False
  txtName.Enabled = False
  txtDarsad.Enabled = False
  
End Sub



Private Sub Form_Load()
   
   frmCodeBimeh.Left = (Screen.Width - frmCodeBimeh.Width) / 2
   frmCodeBimeh.Top = (Screen.Height - frmCodeBimeh.Height) / 2
   'cmdSave.Enabled = False
   
   'Dim d As String
   
   adoCodeBimeh.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + App.Path + "\db1.mdb"
   adoCodeBimeh.RecordSource = "tblKindBimeh"
   adoCodeBimeh.Refresh
   
   con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + App.Path + "\db1.mdb"
   
   txtCodeBimeh.Enabled = False
   txtName.Enabled = False
   txtDarsad.Enabled = False
End Sub
Private Sub ToggleButtons()
   If cmdExit1.Caption = "&Œ—ÊÃ" Then
      cmdExit1.Caption = "&·€Ê"
   Else
      cmdExit1.Caption = "&Œ—ÊÃ"
   End If
   cmdSave.Enabled = Not cmdSave.Enabled
   cmdNew.Enabled = Not cmdNew.Enabled
  'cmdDelete.Enabled = Not cmdDelete.Enabled
   adoCodeBimeh.Enabled = Not adoCodeBimeh.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

