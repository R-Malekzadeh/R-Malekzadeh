VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSodoor 
   Caption         =   "›—„ ’œÊ—ﬁ»÷"
   ClientHeight    =   6555
   ClientLeft      =   1395
   ClientTop       =   1440
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   8760
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
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
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
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtLast 
      Alignment       =   1  'Right Justify
      DataField       =   "Last_Name"
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   2490
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   8415
      Begin MSDataListLib.DataCombo DcmbKindBimeh 
         Bindings        =   "Form1.frx":0000
         DataField       =   "Name"
         DataSource      =   "adoCodeBimeh"
         Height          =   315
         Left            =   5040
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Name"
         BoundColumn     =   "code_Bimeh"
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin VB.TextBox txtDiss 
         Alignment       =   1  'Right Justify
         DataField       =   "Kind_Diss"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtFrist 
         Alignment       =   1  'Right Justify
         DataField       =   "Frist_Name"
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
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DcmbBabat 
         Bindings        =   "Form1.frx":001B
         DataField       =   "Name_Masaref"
         DataSource      =   "adoCodeMasaref"
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Name_Masaref"
         BoundColumn     =   "Code_Masaref"
         Text            =   ""
         RightToLeft     =   -1  'True
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
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
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
         TabIndex        =   17
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblMablag 
         Height          =   375
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
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
         TabIndex        =   15
         Top             =   360
         Width           =   1935
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
         TabIndex        =   14
         Top             =   1920
         Width           =   855
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
         TabIndex        =   13
         Top             =   1320
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
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   135
      Left            =   7200
      TabIndex        =   8
      Top             =   120
      Width           =   15
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   6210
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "12:43 ».Ÿ"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCodeBimeh 
      Height          =   330
      Left            =   120
      Top             =   5280
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
   Begin MSAdodcLib.Adodc adoSodoor 
      Height          =   330
      Left            =   120
      Top             =   5760
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
      Top             =   4800
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
   Begin VB.Label lblTedad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "shom"
      DataSource      =   "adoSodoor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmSodoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim con As New ADODB.Connection





Private Sub cmdDelete_Click()
'   Dim intResponse As Integer
'
'   Beep
'   intResponse = MsgBox("¬Ì«„ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ Õ–› ﬂ‰Ìœ", vbQuestion + vbYesNo, _
'                "Õ–› ﬂ—œ‰")
'   If intResponse = vbYes Then
'      adoSodoor.Recordset.Delete
'   End If
'   Call DataReposition
   adoSodoor.Recordset.Delete adAffectAllChapters
End Sub

Private Sub cmdExit1_Click()
   If cmdExit1.Caption = "&Œ—ÊÃ" Then
      Unload Me
     
   End If
   
   
End Sub

Private Sub cmdNew_Click()
   'Dim Rs As New ADODB.Recordset
   'Dim Con As New ADODB.Connection
   adoSodoor.Recordset.AddNew
   txtFrist.SetFocus
   Call ToggleButtons
   'Rs.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + App.Path + "\db1.mdb"
   'Con.Open "select max(shom) from  tblSodoor ", Con
   'lblTedad = Rs!maxshom + 1
   'Rs.Close
   'Con.Close
   
   
End Sub

Private Sub cmdSave_Click()
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
rs.Open Sql, con
codeb = rs!code_Bimeh
rs.Close

rs.Open "Select * From tblsodoor", con, adOpenKeyset, adLockOptimistic
rs.AddNew
rs!Frist_Name = txtFrist
rs!Last_Name = txtLast
rs!Kind_Diss = txtDiss
rs!Kind_Insurance = codeb
rs!Mablag = Val(lblMablag)
rs!shom = Val(lblTedad)
rs.Update
rs.Close
DcmbBabat = ""
DcmbKindBimeh = ""
lblMablag = ""
txtDiss = ""
txtLast = ""
txtFrist = ""
lblTedad = Val(lblTedad) + 1
txtFrist.SetFocus

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

'Sql = "Select * From tblkindbimeh where name='" & DcmbKindBimeh & "'"
'rs.Open Sql, con
'codeb = rs!Darsad
'n = rs!Name
'rs.Close

sql1 = "Select * From tblmasaref where Name_Masaref='" & DcmbBabat & "'"
rs.Open sql1, con
If rs.EOF = False Then
    codem = rs!Mablag1
End If
rs.Close

Dim n As String

sql2 = "select * from tblKindBimeh where Name='" & DcmbKindBimeh & " ' "
rs.Open sql2, con
n = rs!Name
codeb = Val(rs!Darsad)
If n = "¬“«œ" Then
   lblMablag = codeb * codem
Else
   lblMablag = (codeb * codem) / 100

End If

rs.Close
'a = DcmbBabat.SelectedItem
'Text1.Text = a

End Sub

Private Sub Form_Load()
  frmSodoor.Left = (Screen.Width - frmSodoor.Width) / 2
  frmSodoor.Top = (Screen.Height - frmSodoor.Height) / 2
  Dim h As String
  
  h = App.Path
  
  adoCodeMasaref.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & h & "\DB1.MDB"
  adoCodeMasaref.RecordSource = "tblMasaref"
  adoCodeMasaref.Refresh
  
  adoCodeBimeh.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + h + "\db1.mdb"
  adoCodeBimeh.RecordSource = "tblKindBimeh"
  adoCodeBimeh.Refresh
  
  adoSodoor.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + h + "\db1.mdb"
  adoSodoor.RecordSource = "tblSodoor"
  adoSodoor.Refresh
  
  con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + h + "\db1.mdb"

  
    
  Dim rs As New ADODB.Recordset
   
  rs.Open "select *from tblSodoor", con
  If rs.BOF = True And rs.EOF = True Then
     
     MsgBox "ÃœÊ· ﬁ»Ê÷ Œ«·Ì «”  ", vbSystemModal
        
  End If
  rs.Close
  
  rs.Open "Select max(shom) as mshom from tblsodoor ", con
  If rs.BOF = False And rs.EOF = False Then
     If IsNull(rs!mshom) Then
        lblTedad = 1
     Else
        lblTedad = rs!mshom + 1
     End If
  End If
  rs.Close
  
End Sub
Private Sub DataReposition()
adoSodoor.Recordset.MoveNext
   With adoSodoor.Recordset
      If .EOF Then
         .MovePrevious
      End If
   End With
End Sub
Private Sub ToggleButtons()
   If cmdExit1.Caption = "&Œ—ÊÃ" Then
      cmdExit1.Caption = "&·€Ê"
   Else
      cmdExit1.Caption = "&Œ—ÊÃ"
   End If
   cmdSave.Enabled = Not cmdSave.Enabled
   cmdNew.Enabled = Not cmdNew.Enabled
   cmdDelete.Enabled = Not cmdDelete.Enabled
   adoSodoor.Enabled = Not adoSodoor.Enabled
End Sub


Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub


