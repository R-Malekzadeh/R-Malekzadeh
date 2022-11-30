VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAslah 
   Caption         =   "›—„ «’·«Õ ﬁ»÷"
   ClientHeight    =   7875
   ClientLeft      =   945
   ClientTop       =   570
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   10005
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   3375
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Frist_Name"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Last_Name"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   2490
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   8415
      Begin VB.TextBox txtLast 
         Alignment       =   1  'Right Justify
         DataField       =   "Last_Name"
         DataSource      =   "adoAslah"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "Name_Masaref"
         DataSource      =   "adoAslah"
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
         ItemData        =   "Form3.frx":0017
         Left            =   240
         List            =   "Form3.frx":0021
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "kind_Insurance"
         DataSource      =   "adoAslah"
         Height          =   315
         ItemData        =   "Form3.frx":003A
         Left            =   5040
         List            =   "Form3.frx":0044
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtFrist 
         Alignment       =   1  'Right Justify
         DataField       =   "Frist_Name"
         DataSource      =   "adoAslah"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtDiss 
         Alignment       =   1  'Right Justify
         DataField       =   "Kind_Diss"
         DataSource      =   "adoAslah"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblFrist 
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
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblLast 
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
         Height          =   375
         Left            =   7230
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblBimary 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ »Ì„«—Ì"
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
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblBimeh 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ »Ì„Â"
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
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblMablagMohaseb 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ „Õ«”»Â ‘œÂ «“ »«»  :"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblMablag 
         DataField       =   "mablag"
         DataSource      =   "adoAslah"
         Height          =   375
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblRial 
         Alignment       =   1  'Right Justify
         Caption         =   "—Ì«·"
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ „Õ«”»Â ‘œÂ :"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc adoAslah 
      Height          =   330
      Left            =   120
      Top             =   7440
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
      CommandType     =   8
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
   Begin VB.CommandButton Command1 
      Caption         =   "Ã” ÃÊ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   7080
      Width           =   1395
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
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAslah 
      Caption         =   "«’·«Õ"
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
      Left            =   8040
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoCodeBimeh 
      Height          =   330
      Left            =   120
      Top             =   7080
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc adoCodeMasaref 
      Height          =   330
      Left            =   120
      Top             =   6720
      Visible         =   0   'False
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
End
Attribute VB_Name = "frmAslah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim con1 As New ADODB.Connection




Private Sub cmdAslah_Click()
   Dim rs As New ADODB.Recordset
   
   If Len(txtFrist) = 0 Then
    MsgBox "‰«„ —« Ê«—œ ò‰Ìœ"
    txtFrist.SetFocus
    Exit Sub
End If
    
If Len(txtLast) = 0 Then
    MsgBox "‰«„ Œ«‰Ê«œêÌ —« Ê«—œ ò‰Ìœ"
    txtLast.SetFocus
    Exit Sub
End If

If Len(txtDiss) = 0 Then
    MsgBox "‰Ê⁄ »Ì„«—Ì —« Ê«—œ ò‰Ìœ"
    txtDiss.SetFocus
    Exit Sub
End If

If Len(DcmbKindBimeh) = 0 Then
    MsgBox "‰Ê⁄ »Ì„Â —« Ê«—œ ò‰Ìœ"
    DcmbKindBimeh.SetFocus
    Exit Sub
End If

If Len(DcmbBabat) = 0 Then
    MsgBox "‰Ê⁄ „’—› —« Ê«—œ ò‰Ìœ"
    DcmbBabat.SetFocus
    Exit Sub
End If

Sql = "Select * From tblkindbimeh where name='" & DcmbKindBimeh & "'"
rs.Open Sql, con1
codeb = rs!code_Bimeh
rs.Close

rs.Open "Select * From tblsodoor", con1, adOpenKeyset, adLockOptimistic

rs!Frist_Name = txtFrist
rs!Last_Name = txtLast
rs!Kind_Diss = txtDiss
rs!Kind_Insurance = codeb
rs!Mablag = Val(lblMablag)
rs!shom = Val(lblTedad)
rs.Update
'DataGrid1.Refresh
rs.Close

End Sub

Private Sub Command1_Click()
If txtFrist = "" Then
    adoAslah.RecordSource = "SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name, tblSodoor.Mablag,tblMasaref.Name_Masaref, tblSodoor.Kind_Diss from(tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Kind_Insurance=tblKindBimeh.code_Bimeh)INNER JOIN tblMasaref ON tblSodoor.Babat=tblMasaref.Code_Masaref"
    adoAslah.Refresh
    DataGrid1.ReBind
Else

    adoAslah.RecordSource = "SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name, tblSodoor.Mablag,tblMasaref.Name_Masaref, tblSodoor.Kind_Diss from(tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Kind_Insurance=tblKindBimeh.code_Bimeh)INNER JOIN tblMasaref ON tblSodoor.Babat=tblMasaref.Code_Masaref where tblSodoor.Frist_Name = '" & txtFrist & " ' "
    adoAslah.Refresh
    DataGrid1.ReBind
End If

End Sub


Private Sub DcmbBabat_Click(Area As Integer)
   Dim rs As New ADODB.Recordset
If Len(DcmbKindBimeh) = 0 Then
    MsgBox "‰Ê⁄ »Ì„Â —« Ê«—œ ò‰Ìœ"
    DcmbKindBimeh.SetFocus
    Exit Sub
End If

If Len(DcmbBabat) = 0 Then
    MsgBox "‰Ê⁄ „’—› —« Ê«—œ ò‰Ìœ"
    DcmbBabat.SetFocus
    Exit Sub
End If

sql1 = "Select * From tblmasaref where Name_Masaref='" & DcmbBabat & "'"
rs.Open sql1, con1
If rs.EOF = False Then
    codem = rs!Mablag1
End If
rs.Close

Dim n As String

sql2 = "select * from tblKindBimeh where Name='" & DcmbKindBimeh & "  ' "
rs.Open sql2, con1
n = rs!Name
codeb = Val(rs!Darsad)
If n = "¬“«œ" Then
   lblMablag = codeb * codem
Else
   lblMablag = (codeb * codem) / 100

End If

rs.Close
End Sub

Private Sub DcmbKindBimeh_Click(Area As Integer)
  ' Dim rs1 As New ADODB.Recordset
  ' rs1.Open "Select * from tblKindbimeh", con1
  ' sql = "Select * From tblkindbimeh where name='" & DcmbKindBimeh & "'"
  
  ' rs1.Open sql, con1
  ' code_bimeh = rs1!Name
  ' rs1.Close

  ' DcmbKindBimeh.Tag = DcmbKindBimeh.Text
End Sub


Private Sub Form_Load()
   frmAslah.Left = (Screen.Width - frmAslah.Width) / 2
   frmAslah.Top = (Screen.Height - frmAslah.Height) / 2
  ' Dim d As String

   'd = App.Path
    
   'Adodc1.RecordSource = "SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name,  tblSodoor.Mablag, tblSodoor.Kind_Diss FROM (tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Kind_Insurance = tblKindBimeh.code_Bimeh) INNER JOIN tblMasaref ON (tblKindBimeh.ID = tblMasaref.ID) AND (tblSodoor.Mablag = tblMasaref.Code_Masaref) AND (tblSodoor.Babat = tblMasaref.Code_Masaref)"
   adoAslah.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + App.Path + "\DB1.MDB"
                                                        
   adoAslah.RecordSource = "SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name, tblSodoor.Mablag,tblMasaref.Name_Masaref, tblSodoor.Kind_Diss from(tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Kind_Insurance=tblKindBimeh.code_Bimeh)INNER JOIN tblMasaref ON tblSodoor.Babat=tblMasaref.Code_Masaref"
   
  ' adoAslah.RecordSource = "SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name, tblSodoor.Mablag,tblMasaref.Name_Masaref, tblSodoor.Kind_Diss" _
  '    & "FROM (tblSodoor INNER JOIN (tblKindBimeh INNER JOIN tblMasaref ON tblKindBimeh.ID = tblMasaref.ID) ON (tblSodoor.Babat = tblMasaref.Code_Masaref) AND (tblSodoor.Kind_Insurance = tblKindBimeh.code_Bimeh)"
   '"SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name, tblKindBimeh.Name, tblSodoor.Mablag,tblMasaref.Name_Masaref, tblSodoor.Kind_Diss FROM tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Kind_Insurance = tblKindBimeh.code_Bimeh" & "tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Babat = tblMasaref.Code_Masaref" '"SELECT tblSodoor.shom, tblSodoor.Frist_Name, tblSodoor.Last_Name,tblMasaref.Name_Masaref , tblSodoor.Mablag, tblSodoor.Kind_Diss tblSodoor INNER JOIN tblKindBimeh ON tblSodoor.Babat = tblMasaref.Code_Masaref"
   adoAslah.Refresh
  
   con1.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + App.Path + "\db1.mdb"
  
   ' tblMasaref ON tblKindBimeh.ID = tblMasaref.ID AND
  ' Dim rs As New ADODB.Recordset
 ' rs.Open "select * from tblkindbimeh", con1
  Dim h As String
  h = App.Path
  
  adoCodeMasaref.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & h & "\DB1.MDB"
  adoCodeMasaref.RecordSource = "tblMasaref"
  adoCodeMasaref.Refresh
  
  adoCodeBimeh.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + h + "\db1.mdb"
  adoCodeBimeh.RecordSource = "tblKindBimeh"
  adoCodeBimeh.Refresh


End Sub

Private Sub Form_Unload(Cancel As Integer)
   con1.Close
End Sub


