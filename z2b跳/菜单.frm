VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB版2B跳V1.4 QQ:836613859 ——菜单"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9315
   Icon            =   "菜单.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "菜单.frx":324A
   ScaleHeight     =   7065
   ScaleWidth      =   9315
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   4560
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   3120
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   720
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
      _cy             =   1085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "说  明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "退  出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "开始游戏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8445
      Left            =   0
      Picture         =   "菜单.frx":6494
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()

End Sub

Private Sub Label1_Click()
Form1.Show
'Form2.Hide
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path + "\pic\2yun.jpg")
Picture2.Picture = LoadPicture(App.Path + "\pic\xxoxx.jpg")
'WindowsMediaPlayer1.URL = App.Path + "\BGM\BGM1.mp3"
'WindowsMediaPlayer1.Controls.play
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label3_Click()
Form3.Show
End Sub


