VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00808000&
   Caption         =   "���[�ݤH�ƭp�⤧�h�C��Ʀ�ݪO����t�ι�@"
   ClientHeight    =   7455
   ClientLeft      =   1590
   ClientTop       =   2190
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   12870
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Caption         =   "�d��"
      ForeColor       =   &H0000FFFF&
      Height          =   2775
      Left            =   5160
      TabIndex        =   21
      Top             =   4200
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "�̸`�ر���d��"
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton Output2 
         Caption         =   "�̤�������d��"
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Output1 
         Caption         =   "�̮ɶ�����d��"
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "�`�ت�"
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   2415
      Begin VB.CommandButton OpenViewShow 
         Caption         =   "�w���ýs��Show��ƪ�"
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "����s��"
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   2640
      TabIndex        =   17
      Top             =   4200
      Width           =   2415
      Begin VB.CommandButton OpenData 
         Caption         =   "�w��Record��ƪ�"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame StateFrame 
      BackColor       =   &H00808000&
      Caption         =   "���A"
      ForeColor       =   &H0000FFFF&
      Height          =   7215
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Width           =   4935
      Begin VB.TextBox State 
         Height          =   6855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '�������b
         TabIndex        =   15
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame MainFrame 
      BackColor       =   &H00808000&
      Caption         =   "�ʧ@"
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton OpenAM 
         BackColor       =   &H00000000&
         Caption         =   "�ҰʤH�y�����n��Attention Meter"
         Height          =   735
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton OpenWrite 
         Caption         =   "�Ұʸ�Ƽg�J�{��"
         Height          =   735
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CloseWrite 
         Caption         =   "�פ��Ƽg�J�{��"
         Height          =   735
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame OptionsFrame 
      BackColor       =   &H00808000&
      Caption         =   "�]�w"
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7575
      Begin VB.TextBox SecText 
         Alignment       =   2  '�m�����
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   7335
      End
      Begin VB.TextBox JPsecText 
         Alignment       =   2  '�m�����
         Height          =   270
         Left            =   3000
         TabIndex        =   4
         Text            =   "8"
         Top             =   640
         Width           =   735
      End
      Begin VB.TextBox DLsecText 
         Alignment       =   2  '�m�����
         Height          =   270
         Left            =   3000
         TabIndex        =   3
         Text            =   "67"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox CBsecText 
         Alignment       =   2  '�m�����
         Height          =   270
         Left            =   3000
         TabIndex        =   2
         Text            =   "34"
         Top             =   1040
         Width           =   735
      End
      Begin VB.TextBox SleepmsText 
         Alignment       =   2  '�m�����
         Height          =   270
         Left            =   3000
         TabIndex        =   1
         Text            =   "250"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label SecLabel 
         BackColor       =   &H00808000&
         Caption         =   "�Hface_attention���Ȥ����X�U�d��w���G"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label DLLabel 
         BackColor       =   &H00808000&
         Caption         =   "DL�ԲӬd�ݤ�face_attention�P�_�ȡG"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label CBLabel 
         BackColor       =   &H00808000&
         Caption         =   "CB�H�N�s����face_attention�P�_�ȡG"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1040
         Width           =   2895
      End
      Begin VB.Label JPLabel 
         BackColor       =   &H00808000&
         Caption         =   "JP�u�O���L��face_attention�P�_�ȡG"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   640
         Width           =   2895
      End
      Begin VB.Label SleepmsLebel 
         BackColor       =   &H00808000&
         Caption         =   "�{�ǳB�z���j��ơ]���G�@��^�G"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��Ʈw�w�q
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset
'�ŧiSleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'�{�Ƕ��j�ɶ��Ѽ�
Dim Sleepms As Integer
'�u�O���L���P�_��JP �w�]�� =>8 and <34 JPsec = 8
Dim JPsec As Integer
'�H�N�s�����P�_��CB �w�]�� >34 and <67 CBsec = 34
Dim CBsec As Integer
'�ԲӬd�ݤ��P�_��DL �w�]�� >67 DLsec = 67
Dim DLsec As Integer
'�O�_���b�g�J�P�_
Dim WriteStop As Boolean

Private Sub CBsecText_Change()

SecText.Text = "[ 0��~�P��" & (Val(JPsecText.Text) - 1) & " ][ " & _
JPsecText.Text & "��JP�u�O���L��" & (Val(CBsecText.Text) - 1) & " ][ " & _
CBsecText.Text & "��CB�H�N�s����" & (Val(DLsecText.Text) - 1) & " ][ " & _
DLsecText.Text & "��DL�ԲӬd�ݡ�100]"

State.Text = Now() & "�G" & "�w��s�ﶵ" & vbCrLf & State.Text

End Sub

Private Sub CloseWrite_Click()

State.Text = Now() & "�G" & "���b�פ�g�J�{��" & vbCrLf & State.Text

'�����g�J�{��
WriteStop = True

Set Rec = Nothing
Set Con = Nothing

State.Text = Now() & "�G" & "�w�פ�g�J�{��" & vbCrLf & State.Text

End Sub

Private Sub Command1_Click()

WhereShow.Show

End Sub

Private Sub DLsecText_Change()

SecText.Text = "[ 0��~�P��" & (Val(JPsecText.Text) - 1) & " ][ " & _
JPsecText.Text & "��JP�u�O���L��" & (Val(CBsecText.Text) - 1) & " ][ " & _
CBsecText.Text & "��CB�H�N�s����" & (Val(DLsecText.Text) - 1) & " ][ " & _
DLsecText.Text & "��DL�ԲӬd�ݡ�100]"

State.Text = Now() & "�G" & "�w��s�ﶵ" & vbCrLf & State.Text

End Sub

Private Sub Form_Load()

'�ƭȳ]�w
Sleepms = SleepmsText.Text
JPsec = JPsecText.Text
CBsec = CBsecText.Text
DLsec = DLsecText.Text

State.Text = Now() & "�G" & "��l��"

SecText.Text = "[ 0��~�P��" & (Val(JPsecText.Text) - 1) & " ][ " & _
JPsecText.Text & "��JP�u�O���L��" & (Val(CBsecText.Text) - 1) & " ][ " & _
CBsecText.Text & "��CB�H�N�s����" & (Val(DLsecText.Text) - 1) & " ][ " & _
DLsecText.Text & "��DL�ԲӬd�ݡ�100]"

End Sub


Private Sub JPsecText_Change()

SecText.Text = "[ 0��~�P��" & (Val(JPsecText.Text) - 1) & " ][ " & _
JPsecText.Text & "��JP�u�O���L��" & (Val(CBsecText.Text) - 1) & " ][ " & _
CBsecText.Text & "��CB�H�N�s����" & (Val(DLsecText.Text) - 1) & " ][ " & _
DLsecText.Text & "��DL�ԲӬd�ݡ�100]"

State.Text = Now() & "�G" & "�w��s�ﶵ" & vbCrLf & State.Text

End Sub

Private Sub OpenAM_Click()

State.Text = Now() & "�G" & "���b�Ұ�Attention Meter" & vbCrLf & State.Text

'���}Attention Meter.exe
Dim RetVal
RetVal = Shell("C:\���[�ݤH�ƭp�⤧�h�C��Ʀ�ݪO����t�ι�@\Attention Meter.exe", vbNormalFocus)

State.Text = Now() & "�G" & "�w�Ұ�Attention Meter" & vbCrLf & State.Text

End Sub

Private Sub OpenData_Click()

ViewRecord.Show

End Sub

Private Sub OpenViewShow_Click()

ViewShow.Show

End Sub

Private Sub OpenWrite_Click()

Set Con = Nothing
Set Rec = Nothing

State.Text = Now() & "�G" & "���b�Ұʼg�J�{��" & vbCrLf & State.Text

State.Text = Now() & "�G" & "���bŪ��SQL Server��Ʈw Record��ƪ�" & vbCrLf & State.Text

'�s��SQL Server��Ʈw
'Connection����Con
Con.ConnectionString = "Driver={SQL Server};Server=.;Database=Database;Trusted_Connection=yes;"
Con.Open
'Recordset����Rec
Rec.ActiveConnection = Con
Rec.Open "select * from Record", Con, 1, 3

State.Text = Now() & "�G" & "�w�s����SQL Server��Ʈw Record��ƪ�" & vbCrLf & State.Text

State.Text = Now() & "�G" & "���bŪ���]�w" & vbCrLf & State.Text

Sleepms = SleepmsText.Text
JPsec = JPsecText.Text
CBsec = CBsecText.Text
DLsec = DLsecText.Text

State.Text = Now() & "�G" & "�wŪ���]�w" & vbCrLf & State.Text

State.Text = Now() & "�G" & "�w�Ұʼg�J�{��" & vbCrLf & State.Text

'�ŧi�ܼ�
Dim TempRead
Dim AttentionType As String
Dim TempArrayBefore(500) As String

'��Ʈw��� Number,Time, FaceNumber,AttentionType,Frame, Wx, Wy, AttentionLevel,FaceAmount,
'NoddingAmount, ShakingAmount, MovingAmount, MouthsOpenAmount,X , Y, Width, Height,
'FaceAttention, FaceAge, FaceNodding, FaceShaking, LastBlink, MouthOpen, MouthSmile

'�p��TempArrayBefore�}�C����
UBBefore = UBound(TempArrayBefore)

'�w�]�}�C
For i = 0 To UBBefore - 1
TempArrayBefore(i) = "0"
Next i

'����Ū��attout�����é�����}�CTempArrayNow ���jSleep�@��
'��ƼȦs
Do
Open "C:\���[�ݤH�ƭp�⤧�h�C��Ʀ�ݪO����t�ι�@\attout.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, TempRead
Loop
Close #1

'�r����Ω�JTempArrayNow�}�C
TempArrayNow = Split(TempRead, "&")

'�p��TempArrayNow�}�C����
UBNow = UBound(TempArrayNow)

'�M�����������h�l�r��
TempArrayNow(0) = Replace(TempArrayNow(0), "frame=", "")
TempArrayNow(1) = Replace(TempArrayNow(1), "wx=", "")
TempArrayNow(2) = Replace(TempArrayNow(2), "wy=", "")
TempArrayNow(3) = Replace(TempArrayNow(3), "attentionlevel=", "")
TempArrayNow(4) = Replace(TempArrayNow(4), "face=", "")
TempArrayNow(5) = Replace(TempArrayNow(5), "nodding=", "")
TempArrayNow(6) = Replace(TempArrayNow(6), "shaking=", "")
TempArrayNow(7) = Replace(TempArrayNow(7), "moving=", "")
TempArrayNow(8) = Replace(TempArrayNow(8), "mouthsOpen=", "")
If (Val(TempArrayNow(4)) > 0) Then
For i = 0 To Val(TempArrayNow(4)) - 1
TempArrayNow(9 + (12 * i)) = Replace(TempArrayNow(9 + (12 * i)), "x" & i & "=", "")
TempArrayNow(10 + (12 * i)) = Replace(TempArrayNow(10 + (12 * i)), "y" & i & "=", "")
TempArrayNow(11 + (12 * i)) = Replace(TempArrayNow(11 + (12 * i)), "width" & i & "=", "")
TempArrayNow(12 + (12 * i)) = Replace(TempArrayNow(12 + (12 * i)), "height" & i & "=", "")
TempArrayNow(13 + (12 * i)) = Replace(TempArrayNow(13 + (12 * i)), "face_attention" & i & "=", "")
TempArrayNow(14 + (12 * i)) = Replace(TempArrayNow(14 + (12 * i)), "face_age" & i & "=", "")
TempArrayNow(15 + (12 * i)) = Replace(TempArrayNow(15 + (12 * i)), "face_nodding" & i & "=", "")
TempArrayNow(16 + (12 * i)) = Replace(TempArrayNow(16 + (12 * i)), "face_shaking" & i & "=", "")
TempArrayNow(17 + (12 * i)) = Replace(TempArrayNow(17 + (12 * i)), "face_moving" & i & "=", "")
TempArrayNow(18 + (12 * i)) = Replace(TempArrayNow(18 + (12 * i)), "last_blink" & i & "=", "")
TempArrayNow(19 + (12 * i)) = Replace(TempArrayNow(19 + (12 * i)), "mouthOpen" & i & "=", "")
TempArrayNow(20 + (12 * i)) = Replace(TempArrayNow(20 + (12 * i)), "mouthSmile" & i & "=", "")
Next i
End If

'�}�l�P�_
'�P�_0~n�ӤH�y
For i = 0 To Val(TempArrayNow(4)) - 1

'�P�_TempArrayNow��FaceAttention�O�_�p��TempArrayBefore��FaceAttention
If Val(TempArrayNow(13 + (12 * i))) < Val(TempArrayBefore(13 + (12 * i))) Then

'�p�G�O�h�~��P�_AttentionType �P�_���O�e�@�� �]���H�w�g���F
If TempArrayBefore(13 + ((12 * i))) < JPsec Then
'�~�P
AttentionType = "Null"
ElseIf TempArrayBefore(13 + ((12 * i))) <= CBsec Then
AttentionType = "JP"
ElseIf TempArrayBefore(13 + ((12 * i))) <= DLsec Then
AttentionType = "CB"
ElseIf TempArrayBefore(13 + ((12 * i))) <= 100 Then
AttentionType = "DL"
End If

'�p�G���O�~�P�h�s�i��Ʈw �s�i���O�e�@�� �]���H�w�g���F
If AttentionType <> "Null" Then
Rec.AddNew
Rec.Fields(0) = ""
Rec.Fields(1) = Format(Now(), "yyyy/mm/dd AMPM hh:mm:ss")
Rec.Fields(2) = i
Rec.Fields(3) = AttentionType
Rec.Fields(4) = TempArrayBefore(0)
Rec.Fields(5) = TempArrayBefore(1)
Rec.Fields(6) = TempArrayBefore(2)
Rec.Fields(7) = TempArrayBefore(3)
Rec.Fields(8) = TempArrayBefore(4)
Rec.Fields(9) = TempArrayBefore(5)
Rec.Fields(10) = TempArrayBefore(6)
Rec.Fields(11) = TempArrayBefore(7)
Rec.Fields(12) = TempArrayBefore(8)
Rec.Fields(13) = TempArrayBefore(9 + ((12 * i)))
Rec.Fields(14) = TempArrayBefore(10 + ((12 * i)))
Rec.Fields(15) = TempArrayBefore(11 + ((12 * i)))
Rec.Fields(16) = TempArrayBefore(12 + ((12 * i)))
Rec.Fields(17) = TempArrayBefore(13 + ((12 * i)))
Rec.Fields(18) = TempArrayBefore(14 + ((12 * i)))
Rec.Fields(19) = TempArrayBefore(15 + ((12 * i)))
Rec.Fields(20) = TempArrayBefore(16 + ((12 * i)))
Rec.Fields(21) = TempArrayBefore(17 + ((12 * i)))
Rec.Fields(22) = TempArrayBefore(18 + ((12 * i)))
Rec.Fields(23) = TempArrayBefore(19 + ((12 * i)))
Rec.Fields(24) = TempArrayBefore(20 + ((12 * i)))
Rec.Update

State.Text = Now() & "�G" & "�w�g�J���" & vbCrLf & State.Text

End If
End If
Next i

'TempArrayNow�л\�M�ū᪺TempArrayBefore �M��TempArrayNow�M��
For i = 0 To (UBBefore - 1)
TempArrayBefore(i) = "0"
Next i
For i = 0 To (UBNow - 1)
TempArrayBefore(i) = TempArrayNow(i)
Next i
For i = 0 To (UBNow - 1)
TempArrayNow(i) = "0"
Next i

'�P�_�O�_���b�g�J
If WriteStop Then WriteStop = False: Exit Sub
DoEvents

If WriteStop Then
Else
State.Text = Now() & "�G" & "�{�ǥ��b�i��" & vbCrLf & State.Text
End If

'����
Sleep (Sleepms)

Loop

End Sub


Private Sub Output1_Click()

WhereTime.Show

End Sub

Private Sub Output2_Click()

WhereType.Show

End Sub

Private Sub SleepmsText_Change()

SecText.Text = "[ 0��~�P��" & (Val(JPsecText.Text) - 1) & " ][ " & _
JPsecText.Text & "��JP�u�O���L��" & (Val(CBsecText.Text) - 1) & " ][ " & _
CBsecText.Text & "��CB�H�N�s����" & (Val(DLsecText.Text) - 1) & " ][ " & _
DLsecText.Text & "��DL�ԲӬd�ݡ�100]"

State.Text = Now() & "�G" & "�w��s�ﶵ" & vbCrLf & State.Text

End Sub

