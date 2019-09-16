VERSION 5.00
Begin VB.Form Formsd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登陆及游戏初设定"
   ClientHeight    =   3990
   ClientLeft      =   -735
   ClientTop       =   7125
   ClientWidth     =   9075
   Icon            =   "wzqcj2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wzqcj2.frx":324A
   ScaleHeight     =   3990
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "一键注册"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaxLength       =   4
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定登陆"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FF0000&
      Caption         =   "局域网对战（需要mswinsck.ocx部件，没有的可在网上下载一个）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   7815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808000&
      Caption         =   "人机对战（不要期待有多么的智能）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   5895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800080&
      Caption         =   "单人对战（无聊的自己和自己对弈）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入登录名和密码即可注册"
      ForeColor       =   &H00000040&
      Height          =   360
      Left            =   4920
      TabIndex        =   9
      Top             =   1560
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码：（最长10个字符,只能为数字或英文）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登陆名：(最长4个字符,只能为数字或英文)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5445
   End
End
Attribute VB_Name = "Formsd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim a As dlm, jh%
If Text1 = "" And Text2 = "" Then
   MsgBox "请输入登录名和密码", 48, "提示"
   Text1.SetFocus
   Exit Sub
ElseIf Text1 = "" And Text2 <> "" Then
       MsgBox "请输入登录名", 48, "提示"
       Text1.SetFocus
       Exit Sub
ElseIf Text1 <> "" And Text2 = "" Then
       MsgBox "请输入密码", 48, "提示"
       Text2.SetFocus
       Exit Sub
End If
Open App.Path & "zcb.lsn" For Random As #2 Len = Len(a)
     If LOF(2) - 174 <= 0 Then
        MsgBox "请先注册", 48, "提示"
     End If
     For i = 1 To LOF(2) / Len(a) - 1
         Get #2, i + 1, a
         If Trim(a.mz) = Trim(Text1) And Trim(a.mm) = Trim(Text2) Then
            dl.mz = a.mz
            jh = 1
            Exit For
         ElseIf Trim(a.mz) = Trim(Text1) And Trim(a.mm) <> Trim(Text2) Then
                jh = 2
                Exit For
         ElseIf Trim(a.mz) <> Trim(Text1) Then
                jh = 3
         End If
     Next i
Close #2
If jh = 1 Then
   If Option1.Value = True Then
               md = 1
               Unload Me
               Formzjm.Show
   ElseIf Option2 = True Then
                   md = 2
                   Unload Me
                   Formzjm.Show
   ElseIf Option3 = True Then
                   On Error GoTo errwz
                   md = 3
                   Unload Me
                   Formzjm.Show
errwz:
                   If Err.Number = "399" Then
                      MsgBox "电脑无mswinsck.ocx", 48, "提示"
                   Else
                       On Error Resume Next
                   End If
   End If
ElseIf jh = 2 Then
       MsgBox "密码输入错误，请重新输入", 48, "提示"
       Text2 = ""
       Text2.SetFocus
ElseIf jh = 3 Then
       MsgBox "登录名输入有误，请重新输入", 48, "提示"
       Text1 = "": Text2 = ""
       Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Dim fl As Boolean
fl = True
If Text1 = "" And Text2 = "" Then
   MsgBox "请输入登录名和密码", 48, "提示"
   Text1.SetFocus
   Exit Sub
ElseIf Text1 = "" And Text2 <> "" Then
       MsgBox "请输入登录名", 48, "提示"
       Text1.SetFocus
       Exit Sub
ElseIf Text1 <> "" And Text2 = "" Then
       MsgBox "请输入密码", 48, "提示"
       Text2.SetFocus
       Exit Sub
End If
Open App.Path & "zcb.lsn" For Random As #1 Len = Len(dl)
     For i = 1 To LOF(1) / Len(dl) - 1
         Get #1, i + 1, dl
         If Trim(dl.mz) = Trim(Text1) Then
            fl = False
            Exit For
         End If
     Next i
     If fl = True Then
        dl.mz = Text1: dl.mm = Text2
        If LOF(1) = 0 Then
           i = 2
        Else
            i = LOF(1) / Len(dl) + 1
        End If
        Put #1, i, dl
        MsgBox "注册成功", 48, "提示"
     End If
Close #1
If fl = False Then
   MsgBox "此登录名已存在", 48, "提示"
End If
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Dim ran!(1 To 8)
Randomize
For i = 1 To 8
    ran(i) = Int(Rnd * (RGB(255, 255, 255) + 1))
Next i
Command1.BackColor = ran(1)
Command2.BackColor = ran(2)
Option1.BackColor = ran(3): Option1.ForeColor = ran(6)
Option2.BackColor = ran(4): Option2.ForeColor = ran(7)
Option3.BackColor = ran(5): Option3.ForeColor = ran(8)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
For i = 1 To Len(Text1)
    If (Asc(Mid(Text1, i, 1)) > 47 And Asc(Mid(Text1, i, 1)) < 58) Or _
    (Asc(Mid(Text1, i, 1)) > 64 And Asc(Mid(Text1, i, 1)) < 91) Or _
    (Asc(Mid(Text1, i, 1)) > 96 And Asc(Mid(Text1, i, 1)) < 123) Then
    Else
        Text1 = ""
        Text1.SetFocus
        Exit Sub
    End If
Next i
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
For i = 1 To Len(Text2)
    If (Asc(Mid(Text2, i, 1)) > 47 And Asc(Mid(Text2, i, 1)) < 58) Or _
    (Asc(Mid(Text2, i, 1)) > 64 And Asc(Mid(Text2, i, 1)) < 91) Or _
    (Asc(Mid(Text2, i, 1)) > 96 And Asc(Mid(Text2, i, 1)) < 123) Then
    Else
        Text2 = ""
        Text2.SetFocus
        Exit Sub
    End If
Next i
End Sub
