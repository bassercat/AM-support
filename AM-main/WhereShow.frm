VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form WhereShow 
   BackColor       =   &H00808000&
   Caption         =   "WhereShow"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   15615
   StartUpPosition =   3  '系統預設值
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6735
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   11880
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
   Begin VB.ComboBox showbox 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "WhereShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'資料庫定義
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset

Private Sub Command1_Click()

Dim APM1 As String
Dim APM2 As String
Dim Time1 As String
Dim Time2 As String

Set Con = Nothing
Set Rec = Nothing

'連結SQL Server資料庫
'Connection物件Con
Con.ConnectionString = "Driver={SQL Server};Server=.;Database=Database;Trusted_Connection=yes;"
Con.Open
'Recordset物件Rec
Rec.ActiveConnection = Con

TempShow = Split(showbox.Text, " / ")
TempTime = Split(TempShow(1), " ~ ")
Time1 = TempTime(0)
Time2 = TempTime(1)
If (Left(Time1, 2) > 12) Then
APM1 = "下午"
APM2 = "下午"
Time1 = Replace(Time1, Val(Left(Time1, 2)), Val(Left(Time1, 2)) - 12)
Time2 = Replace(Time2, Val(Left(Time2, 2)), Val(Left(Time2, 2)) - 12)
ElseIf (Left(Time2, 2) > 12) Then
APM1 = "上午"
APM2 = "下午"
Else
APM1 = "上午"
APM2 = "上午"
End If

Rec.CursorLocation = adUseClient

Rec.Open "select * from Record where SUBSTRING(time, 12, 11) between '" & APM1 & " " & Time1 & "' and '" & APM2 & " " & Time2 & "' order by Time ASC"

Label1.Caption = "select * from Record where SUBSTRING(time, 12, 11) between '" & APM1 & " " & Time1 & "' and '" & APM2 & " " & Time2 & "' order by Time ASC"

Set DataGrid1.DataSource = Rec
DataGrid1.Refresh

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim Temp As String

'連結SQL Server資料庫
'Connection物件Con
Con.ConnectionString = "Driver={SQL Server};Server=.;Database=Database;Trusted_Connection=yes;"
Con.Open
'Recordset物件Rec
Rec.ActiveConnection = Con
Rec.Open "select * from Show", Con, 1, 3

While Not Rec.EOF
showbox.AddItem Rec.Fields(1) & " / " & Rec.Fields(2) & " ~ " & Rec.Fields(3)
Rec.MoveNext
Wend
showbox.Text = showbox.List(0)

End Sub
