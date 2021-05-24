VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form WhereTime 
   BackColor       =   &H00808000&
   Caption         =   "WhereTime"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   14190
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   375
      Left            =   960
      TabIndex        =   31
      Top             =   1440
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   30
      Top             =   2040
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8705
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
   Begin VB.CommandButton Where 
      Caption         =   "查詢"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox Sec2 
      Height          =   300
      Left            =   9960
      TabIndex        =   26
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Min2 
      Height          =   300
      Left            =   8310
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox APM2 
      Height          =   300
      Left            =   5040
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Hour2 
      Height          =   300
      Left            =   6675
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Sec1 
      Height          =   300
      Left            =   9960
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Min1 
      Height          =   300
      Left            =   8310
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox APM1 
      Height          =   300
      Left            =   5040
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Hour1 
      Height          =   300
      Left            =   6675
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Day2 
      Height          =   300
      Left            =   3390
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Year2 
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Month2 
      Height          =   300
      Left            =   1755
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Day1 
      Height          =   300
      Left            =   3390
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Month1 
      Height          =   300
      Left            =   1755
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Year1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Labelx 
      BackColor       =   &H00808000&
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   32
      Top             =   1440
      Width           =   9735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808000&
      Caption         =   "至"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808000&
      Caption         =   "秒"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11280
      TabIndex        =   27
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808000&
      Caption         =   "分"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   25
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808000&
      Caption         =   "時"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   24
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808000&
      Caption         =   "午"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808000&
      Caption         =   "秒"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11280
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808000&
      Caption         =   "分"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808000&
      Caption         =   "時"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808000&
      Caption         =   "午"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      Caption         =   "日"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "月"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "年"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "日"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "月"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "年"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "WhereTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'資料庫定義
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

For i = 2013 To 2015
Year1.AddItem i
Year2.AddItem i
Next i
For i = 1 To 12
If i < 10 Then
Month1.AddItem "0" & i
Month2.AddItem "0" & i
Else
Month1.AddItem i
Month2.AddItem i
End If
Next i
For i = 1 To 31
If i < 10 Then
Day1.AddItem "0" & i
Day2.AddItem "0" & i
Else
Day1.AddItem i
Day2.AddItem i
End If
Next i
APM1.AddItem "上午"
APM1.AddItem "下午"
APM2.AddItem "上午"
APM2.AddItem "下午"
For i = 0 To 12
If i < 10 Then
Hour1.AddItem "0" & i
Hour2.AddItem "0" & i
Else
Hour1.AddItem i
Hour2.AddItem i
End If
Next i
For i = 0 To 60
If i < 10 Then
Min1.AddItem "0" & i
Min2.AddItem "0" & i
Else
Min1.AddItem i
Min2.AddItem i
End If
Next i
For i = 0 To 60
If i < 10 Then
Sec1.AddItem "0" & i
Sec2.AddItem "0" & i
Else
Sec1.AddItem i
Sec2.AddItem i
End If
Next i

Dim Com As ADODB.Command
Dim param1 As ADODB.Parameter
Dim param2 As ADODB.Parameter
Dim ret1 As String
Dim ret2 As String
Dim mySQL As String

Year1.Text = Year1.List(0)
Year2.Text = Year2.List(0)
Month1.Text = Month1.List(0)
Month2.Text = Month2.List(0)
Day1.Text = Day1.List(0)
Day2.Text = Day2.List(0)
APM1.Text = APM1.List(0)
APM2.Text = APM2.List(0)
Hour1.Text = Hour1.List(0)
Hour2.Text = Hour2.List(0)
Min1.Text = Min1.List(0)
Min2.Text = Min2.List(0)
Sec1.Text = Sec1.List(0)
Sec2.Text = Sec2.List(0)

End Sub


Private Sub Where_Click()

Set Con = Nothing
Set Rec = Nothing

Dim Date1 As String
Dim Date2 As String
Date1 = Year1.Text & "/" & Month1.Text & "/" & Day1.Text & " " & APM1.Text & " " & Hour1.Text & ":" & Min1.Text & ":" & Sec1.Text
Date2 = Year2.Text & "/" & Month2.Text & "/" & Day2.Text & " " & APM2.Text & " " & Hour2.Text & ":" & Min2.Text & ":" & Sec2.Text

'連結SQL Server資料庫
'Connection物件Con
Con.ConnectionString = "Driver={SQL Server};Server=.;Database=Database;Trusted_Connection=yes;"
Con.Open
'Recordset物件Rec
Rec.ActiveConnection = Con
Rec.Open "select * from Record where Time between '" & Date1 & "' and '" & Date2 & "' order by Time ASC", Con, 1, 3
Labelx.Caption = "select * from Record where Time between '" & Date1 & "' and '" & Date2 & "' order by Time ASC"

Set DataGrid1.DataSource = Rec
DataGrid1.Refresh

End Sub
