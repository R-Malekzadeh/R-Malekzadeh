VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCodeMasaref 
   BackColor       =   &H00C0C0C0&
   Caption         =   "›—„ ﬂœ „’«—›"
   ClientHeight    =   6030
   ClientLeft      =   1395
   ClientTop       =   1710
   ClientWidth     =   8430
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
   ScaleHeight     =   6030
   ScaleWidth      =   8430
   Begin MSAdodcLib.Adodc adoMasaref 
      Height          =   330
      Left            =   120
      Top             =   5640
      Visible         =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CodeMasaref.frx":0000
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "„»·€ „Õ«”»Â ‘œÂ «“ »«» "
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Code_Masaref"
         Caption         =   "ﬂœ „’«—›"
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
         DataField       =   "Name_Masaref"
         Caption         =   "‰«„ „’«—›"
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
         DataField       =   "Mablag1"
         Caption         =   "„»·€"
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
      Height          =   1455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtMasaref 
         Alignment       =   1  'Right Justify
         DataField       =   "Code_Masaref"
         DataSource      =   "adoMasaref"
         Height          =   360
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtNameMasaref 
         Alignment       =   1  'Right Justify
         DataField       =   "Name_Masaref"
         DataSource      =   "adoMasaref"
         Height          =   360
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtDarsad 
         Alignment       =   1  'Right Justify
         DataField       =   "Mablag1"
         DataSource      =   "adoMasaref"
         Height          =   330
         Left            =   915
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   2250
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ „’«—›"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ „’«—›"
         Height          =   255
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "„»·€"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCodeMasaref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim con As New ADODB.Connection
Private Sub Form_Load()
   frmMain.Left = (Screen.Width - frmMain.Width) / 2
   frmMain.Top = (Screen.Height - frmMain.Height) / 2
   
   adoMasaref.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\DB1.MDB"
   adoMasaref.RecordSource = "tblMasaref"
   adoMasaref.Refresh

   con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\DB1.MDB"

End Sub


Private Sub Form_Unload(Cancel As Integer)
   con.Close
End Sub
