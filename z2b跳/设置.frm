VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "设置"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   3540
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   720
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   4
      Text            =   "495"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Text            =   "10"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "台阶长度（255-999）："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "人物移动速度（游戏中的灵敏度 7-14）："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1) > 2 Then
MsgBox "人物移动速度要为整数！"
GoTo ex:
End If
If Len(Text2) <> 3 Then
MsgBox "台阶长度要为整数！"
GoTo ex:
End If

If Text1 < 7 Or Text1 > 14 Or Text2 < 255 Or Text2 > 999 Then
MsgBox "参数错误！", vbCritical, "错误"
GoTo ex:
End If
Form1.Text2 = Text1
For i = 0 To 10
Form1.Picture2(i).Width = Text2
Next
Kill App.Path + "\setting.ini"
Open App.Path + "\setting.ini" For Append As #1
Print #1, "[SETTING]"
Print #1, "MouseSp=" & Text1
Print #1, "SWidth=" & Text2
Close #1
MsgBox "保存完成！请自行关闭“设置”窗口"
ex:
End Sub

Private Sub Form_Load()
Open App.Path + "\setting.ini" For Input As #1
Do While Not EOF(1)
Line Input #1, a
w = w + a
Loop
msp = Mid(w, InStr(1, w, "MouseSp=") + 8, 2)
wid = Mid(w, InStr(1, w, "SWidth=") + 7, 3)
Close #1
Text1 = msp
Text2 = wid
End Sub

Private Sub Text2_Change()
On Error Resume Next
Picture1.Width = Text2
End Sub
