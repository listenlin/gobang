VERSION 5.00
Begin VB.Form Formhy 
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   3825
   ClientWidth     =   8970
   Icon            =   "wzqcj1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   0
      Picture         =   "wzqcj1.frx":324A
      ScaleHeight     =   4575
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "下 次 进 入 游 戏 不 再 显 示 此 界 面"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   4080
         Width           =   4695
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   7680
         Top             =   2280
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "本游戏由森哥哥独立设计，耗尽脑汁，历时四个月，用VB编制而成！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2895
         Left            =   1800
         TabIndex        =   5
         Top             =   2520
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎来到森林五子棋"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   720
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   6480
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   7680
         Picture         =   "wzqcj1.frx":224E3
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "本游戏由森哥哥独立设计，耗尽脑汁，历时四个月，用VB编制而成！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   2895
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎来到森林五子棋"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   720
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   6480
      End
   End
End
Attribute VB_Name = "Formhy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Timer1.Enabled = True
Label1.Visible = False
Label2.Visible = False
Label4.Visible = False
Label5.Visible = False
Label1.Left = (Me.Width - Label1.Width) / 2
Label1.Top = (Me.Height - Label1.Height) / 2
Label4.Top = Label1.Top + 52.5
Label4.Left = Label1.Left + 52.5
Label2.Left = (Me.Width - Label2.Width) / 2
Label2.Top = (Me.Height - Label2.Height) / 2
Label5.Top = Label2.Top + 52.5
Label5.Left = Label2.Left + 52.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hy As dlm
Open App.Path & "zcb.lsn" For Random As #1 Len = Len(hy)
     If Check1.Value = 1 Then
        hy.mz = "nocx"
        Put #1, 1, hy
     End If
Close #1
End Sub

Private Sub Picture1_Click()
Unload Me
Formsd.Show
End Sub

Private Sub Timer1_Timer()
Label1.Visible = True
Label4.Visible = True
Label1.Top = Label1.Top - 50
Label4.Top = Label4.Top - 50
If Label4.Top < (Me.Height - Label4.Height) / 3 Then
   Label2.Top = Label1.Top + Label1.Height + 400
   Label5.Top = Label2.Top + 52.5
   Label2.Visible = True
   Label5.Visible = True
   Label2.Top = Label2.Top - 50
   Label5.Top = Label5.Top - 50
End If
If Label2.Top > 200 And Label2.Top < 260 Then
   Unload Me
   Formsd.Show
End If
End Sub
