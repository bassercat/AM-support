VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form WhereType 
   BackColor       =   &H00808000&
   Caption         =   "WhereType"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11760
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
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
            LCID            =   1028
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
            LCID            =   1028
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
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox Type1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "WhereType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'資料庫定義
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset

Private Sub Command1_Click()

Set Con = Nothing
Set Rec = Nothing

'連結SQL Server資料庫
'Connection物件Con
Con.ConnectionString = "Driver={SQL Server};Server=.;Database=Database;Trusted_Connection=yes;"
Con.Open
'Recordset物件Rec
Rec.ActiveConnection = Con

If Type1.Text = "只是路過JP" Then
Rec.Open "select * from Record where AttentionType like '%JP%' order by Time ASC", Con, 1, 3
Label1.Caption = "select * from Record where AttentionType like '%JP%' order by Time ASC"
ElseIf Type1.Text = "隨意瀏覽CB" Then
Rec.Open "select * from Record where AttentionType like '%CB%' order by Time ASC", Con, 1, 3
Label1.Caption = "select * from Record where AttentionType like '%CB%' order by Time ASC"
ElseIf Type1.Text = "詳細查看DL" Then
Rec.Open "select * from Record where AttentionType like '%DL%' order by Time ASC", Con, 1, 3
Label1.Caption = "select * from Record where AttentionType like '%DL%' order by Time ASC"
End If
Set DataGrid1.DataSource = Rec
DataGrid1.Refresh

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

Type1.AddItem "只是路過JP"
Type1.AddItem "隨意瀏覽CB"
Type1.AddItem "詳細查看DL"
Type1.Text = Type1.List(0)

End Sub

