VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2B跳 V1.4——游戏界面"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8940
   Icon            =   "2b跳.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8940
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox label4 
      BackColor       =   &H80000007&
      Height          =   425
      Left            =   10440
      Picture         =   "2b跳.frx":324A
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   20
      Top             =   7320
      Width           =   425
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   7440
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H8000000C&
      Height          =   270
      Left            =   4080
      TabIndex        =   19
      Text            =   "密码"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   600
      Top             =   7440
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   4320
      TabIndex        =   16
      Text            =   "10"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   7440
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   120
      Top             =   6960
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   5160
      Picture         =   "2b跳.frx":3668
      ScaleHeight     =   195
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      Picture         =   "2b跳.frx":3C6D
      ScaleHeight     =   345
      ScaleWidth      =   150
      TabIndex        =   7
      Top             =   3720
      Width           =   175
   End
   Begin VB.Timer Timer6 
      Interval        =   50
      Left            =   120
      Top             =   6480
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   390
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   9
      Left            =   7680
      Picture         =   "2b跳.frx":4129
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   8
      Left            =   3000
      Picture         =   "2b跳.frx":4680
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Timer Timer5 
      Interval        =   30
      Left            =   120
      Top             =   6000
   End
   Begin VB.Timer Timer4 
      Interval        =   50
      Left            =   120
      Top             =   5520
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   7
      Left            =   6120
      Picture         =   "2b跳.frx":4BD2
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   6
      Left            =   4080
      Picture         =   "2b跳.frx":5129
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   6960
      Picture         =   "2b跳.frx":5680
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2760
      Picture         =   "2b跳.frx":5BD2
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   4200
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   4440
      Picture         =   "2b跳.frx":6124
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   6120
      Picture         =   "2b跳.frx":6676
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   3720
      Picture         =   "2b跳.frx":6BCD
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   6360
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   120
      Top             =   5040
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   5040
      Picture         =   "2b跳.frx":7124
      ScaleHeight     =   75
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   7680
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   120
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   120
      Top             =   3960
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP2 
      Height          =   495
      Left            =   2880
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1720
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1508
      _cy             =   873
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2280
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "重新开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "开挂："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFC0&
      X1              =   0
      X2              =   1200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFC0&
      X1              =   0
      X2              =   1320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFC0&
      X1              =   -120
      X2              =   1320
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "分数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -3120
      Picture         =   "2b跳.frx":767B
      Top             =   -960
      Width           =   15360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '鼠标控制2B的声明
Private Type POINTAPI
x As Long
y As Long
End Type
Dim z As POINTAPI
Dim n, n2, c, BGMT, sps, vsb, wlgq


Private Sub Form_Load() '窗体加载&背景音乐开始
Picture1.Picture = LoadPicture(App.Path + "/pic/man.jpg")
Open App.Path + "\setting.ini" For Input As #1
Do While Not EOF(1)
Line Input #1, a
w = w + a
Loop
msp = Mid(w, InStr(1, w, "MouseSp=") + 8, 2)
wid = Mid(w, InStr(1, w, "SWidth=") + 7, 3)
Close #1
Text2 = msp
For i = 0 To 10
Picture2(i).Width = wid
Next
wlgq = 0
vsb = 2
n = 0
n2 = 20
c = 0
BGMT = 0
WindowsMediaPlayer1.URL = App.Path + "/BGM/BGM2.MP3"
WindowsMediaPlayer1.Controls.play
End Sub







Private Sub Image1_Click()

End Sub

Private Sub Label2_Click() '重新开始
wlgq = 0
Text1 = "0"
Picture1.Left = 720
Picture1.Top = 3000
label4.Left = 20000
label4.Top = 20000
Text3.Enabled = True
For i = 1 To 10
R = Fix(Rnd * (2200) + 2200)
R2 = Fix(Rnd * (R))
Picture2(i).Left = R * 2
Picture2(i).Top = R2 * 2.2
Picture2(i).Visible = True
Next
Picture2(7).Left = 720
Picture2(7).Top = 6000
Picture2(8).Left = 720 + Picture2(7).Width
Picture2(8).Top = 6000
Picture2(9).Left = Picture2(8).Width + Picture2(8).Left
Picture2(9).Top = 6000

Timer8.Enabled = False
Timer9.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label5_Click()
Form4.Show
End Sub



Private Sub Text1_Change()
If Text1 = 1800 Or Text1 = 6000 Or Text1 = 9600 Or Text1 = 15000 Or Text1 = 21000 Or Text1 = 28000 Or Text1 = 35000 Then
label4.Left = 1000
label4.Top = 100
End If
If Text1 = 23400 Or Text1 = 47400 Or Text1 = 71400 Or Text1 = 119400 Then Label7.Visible = True
If Text1 = 24000 Or Text1 = 48000 Or Text1 = 72000 Or Text1 = 120000 Then
monster.Enabled = True
Label7.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
Text3.Enabled = False
Label2.Enabled = False
monster.Enabled = True
monster.Show
res = Fix(Rnd * (Right(Time, 2)) + 1) Mod 3
lie_truth = Right(Time, 1) Mod 2

If res = 0 And lie_truth = 1 Then monster.Label1 = "跟你说，我这次肯定出 菠萝，你出什么？"
If res = 0 And lie_truth = 0 Then monster.Label1 = "跟你说，我这次肯定出 胶钳，你出什么？"

If res = 1 And lie_truth = 1 Then monster.Label1 = "跟你说，我这次肯定出 胶钳，你出什么？"
If res = 1 And lie_truth = 0 Then monster.Label1 = "跟你说，我这次肯定出 叉烧包，你出什么？"

If res = 2 And lie_truth = 1 Then monster.Label1 = "跟你说，我这次肯定出 叉烧包，你出什么？"
If res = 2 And lie_truth = 0 Then monster.Label1 = "跟你说，我这次肯定出 菠萝，你出什么？"
End If
End Sub

Private Sub Text3_Change()
If wlgq = 5 Then GoTo ex:
If Text3 = "我了个去" Then
Text3 = ""
Timer10.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
wlgq = wlgq + 1
End If
ex:
End Sub

Private Sub Text3_Click()
Text3 = ""
End Sub

Private Sub Timer1_Timer() '鼠标控制2B
On Error Resume Next
GetCursorPos z
Picture1.Left = z.x * Text2 Mod Form1.Width

End Sub

Private Sub Timer10_Timer()
sps = sps + 4
Picture1.Top = Picture1.Top - sps
If sps >= 300 Then
sps = 0
Timer2.Enabled = True
Timer10.Enabled = False
End If
End Sub

Private Sub Timer2_Timer() '坠落&判断失败
n = n + 0.5
If n >= 100 Then n = 100
Picture1.Top = Picture1.Top + n ^ 2
If Picture1.Top > Form1.Height Then
n = 0
WMP2.URL = App.Path + "/Snd/a.mp3"
WMP2.Controls.play
MsgBox "2B死了。。。", vbCritical
Text3.Enabled = False
Timer1.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer9.Enabled = False
Timer2.Enabled = False
End If

If Picture1.Top - label4.Top <= 300 And Picture1.Top + 375 >= label4.Top And Picture1.Left >= Label1.Left And Picture1.Left + Picture1.Width <= label4.Left + label4.Width Then
n = 0
label4.Left = 20000
label4.Top = 20000
Timer8.Enabled = True
Timer3.Enabled = False
Timer2.Enabled = False
End If

For i = 0 To 10
If Picture1.Top + 150 <= Picture2(i).Top + Picture2(i).Height And Picture1.Top + 375 >= Picture2(i).Top And Picture1.Left + 130 > Picture2(i).Left And Picture1.Left < Picture2(i).Left + Picture2(i).Width Then
a = "T"
n = 0
End If
Next
If Picture1.Top + 150 <= Picture2(10).Top + Picture2(10).Height And Picture1.Top + 375 >= Picture2(10).Top And Picture1.Left + 130 > Picture2(10).Left And Picture1.Left < Picture2(10).Left + Picture2(10).Width Then n2 = 27
If a = "T" Then
Timer3.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer() '跳跃

If Picture1.Top - label4.Top <= 300 And Picture1.Top + 375 >= label4.Top And Picture1.Left >= Label1.Left And Picture1.Left + Picture1.Width <= label4.Left + label4.Width Then
n = 0
label4.Left = 20000
label4.Top = 20000
Timer8.Enabled = True
Timer3.Enabled = False
Timer2.Enabled = False
End If

n2 = n2 - 1
Picture1.Top = Picture1.Top - n2 ^ 2
If n2 = 0 Then
Timer2.Enabled = True
n2 = 20
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer() '抽取新台阶坐标
On Error Resume Next
If Text1 > 10000 Then
For i = 0 To 4
If Picture2(i).Top >= Form1.Height Then
R = Fix(Rnd * (3200) + 1000)
R2 = Fix(Rnd * (Form1.Width))
Picture2(i).Left = R2
Picture2(i).Top = R
End If
Next
End If

For i = 0 To 10
If Picture2(i).Top >= Form1.Height Then
R = Fix(Rnd * (3200) + 1000)
R2 = Fix(Rnd * (Form1.Width))
Picture2(i).Left = R2
Picture2(i).Top = R
End If
Next
ex:
End Sub

Private Sub Timer5_Timer() '全台阶向下移动
On Error Resume Next
If Picture1.Top < 3800 Then
For i = 0 To 10
Picture2(i).Top = Picture2(i).Top + 100
Next
Picture1.Top = Picture1.Top + 100
label4.Top = label4.Top + 100
Text1 = Text1 + 12
End If
End Sub

Private Sub Timer6_Timer() '移动台阶
c = c + 1
Picture2(0).Left = c * 100
Picture2(1).Left = 10800 - c * 100

If Text1 > 10000 Then
Picture2(2).Left = c * 120
Picture2(3).Left = 12000 - c * 120
Picture2(4).Left = 13000 - c * 130
End If
If c = 107 Then c = 0

End Sub

Private Sub Timer7_Timer() '背景音乐
BGMT = BGMT + 1
If BGMT = 105 Then
BGMT = 0
WindowsMediaPlayer1.URL = App.Path + "/BGM/BGM2.MP3"
WindowsMediaPlayer1.Controls.play
End If
If Text1 > 22000 Then Call all_visible
End Sub

Private Sub Timer8_Timer() '火箭
sps = sps + 3
Picture1.Top = Picture1.Top - sps
If sps >= 350 Then
sps = 0
Timer2.Enabled = True
Timer8.Enabled = False
End If
End Sub

Private Sub Timer9_Timer() '闪烁台阶
If Text1 >= 22000 Then
Call all_visible
Timer9.Enabled = False
End If

If Text1 <= 18000 Then GoTo ex:
vsb = vsb + 1
For v = 0 To 10
Picture2(v).Visible = vsb Mod 2
Next
ex:
End Sub
Private Sub all_visible()
For i = 0 To 10
Picture2(i).Visible = True
Next
End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub

Private Sub WMP2_OpenStateChange(ByVal NewState As Long)

End Sub
