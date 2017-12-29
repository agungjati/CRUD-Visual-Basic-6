VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   11
      Text            =   "Cari berdasar nama ... "
      Top             =   6360
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operator"
      Height          =   2655
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Text            =   "Umur"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "Form1.frx":0000
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Text            =   "Nama"
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000004&
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\MemberDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\MemberDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Member"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0007
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Klik tabel untuk ke menu Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "data masih ada yang kurang"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!Nama = Text1
Adodc1.Recordset!Alamat = Text2
Adodc1.Recordset!Umur = Text3 'nomor
Adodc1.Recordset.Update
Adodc1.Refresh
End If
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "data masih ada yang kurang"
Else
Adodc1.Recordset!Nama = Text1
Adodc1.Recordset!Alamat = Text2
Adodc1.Recordset!Umur = Text3 'nomor
Adodc1.Recordset.Update
Adodc1.Refresh
Command2.Visible = False
Command5.Visible = False
End If
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
Adodc1.Refresh
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Command1.Visible = True
Command2.Visible = False
Command5.Visible = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub DataGrid1_Click()
If Text1.Text = Null Or Text2.Text = Null Or Text3.Text = Null Then
MsgBox ("Ada yang kosong")
Else
Text1.Text = Adodc1.Recordset!Nama
Text2.Text = Adodc1.Recordset!Alamat
Text3.Text = Adodc1.Recordset!Umur
Command2.Visible = True
Command5.Visible = True
End If
End Sub

Private Sub Text4_Change()
If Text4.Text = "" Then
Adodc1.RecordSource = "select * from member"
Adodc1.Refresh
Else
Adodc1.RecordSource = "select * from member where nama='" & Text3.Text & "'"
Adodc1.Refresh
End If
End Sub
