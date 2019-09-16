VERSION 5.00
Begin VB.Form Formys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选取颜色"
   ClientHeight    =   3345
   ClientLeft      =   -15
   ClientTop       =   1110
   ClientWidth     =   6390
   ForeColor       =   &H00800000&
   Icon            =   "wzqcj3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wzqcj3.frx":324A
   ScaleHeight     =   3345
   ScaleWidth      =   6390
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取    消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "选择此方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "选择此方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "选择此方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      ItemData        =   "wzqcj3.frx":BCDA
      Left            =   2160
      List            =   "wzqcj3.frx":BCDC
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确   定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Textb 
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
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Textg 
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
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Textr 
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
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "高级：输入十六进制数或少于16777215的数设置颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   840
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "快捷选择颜色："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   210
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1920
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自己调色输入0-255之间的数值："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   420
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   2085
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B(蓝色值)"
      Height          =   180
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G(绿色值)"
      ForeColor       =   &H00004000&
      Height          =   180
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R(红色值)"
      ForeColor       =   &H00800080&
      Height          =   180
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   810
   End
End
Attribute VB_Name = "Formys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option2.Value = True Then
   If List1.Selected(0) = True Then ysz = &H0&
   If List1.Selected(1) = True Then ysz = &HFF&
   If List1.Selected(2) = True Then ysz = &HFF00&
   If List1.Selected(3) = True Then ysz = &HFFFF&
   If List1.Selected(4) = True Then ysz = &HFF0000
   If List1.Selected(5) = True Then ysz = &HFF00FF
   If List1.Selected(6) = True Then ysz = &HFFFF00
   If List1.Selected(7) = True Then ysz = &HFFFFFF
   If List1.Selected(8) = True Then ysz = &H228B22
   If List1.Selected(9) = True Then ysz = &HC0C0C0
   If yscd = 1 Then
      ys1 = ysz
   ElseIf yscd = 2 Then
          ys2 = ysz
   ElseIf yscd = 3 Then
          ys3 = ysz
   ElseIf yscd = 4 Then
          ys4 = ysz
   End If
   cdgb = False
   Unload Me
ElseIf Option3.Value = True Then
       ysz = RGB(Val(Textr), Val(Textg), Val(Textb))
       If yscd = 1 Then
          ys1 = ysz
       ElseIf yscd = 2 Then
              ys2 = ysz
       ElseIf yscd = 3 Then
              ys3 = ysz
       ElseIf yscd = 4 Then
          ys4 = ysz
       End If
       cdgb = False
       Unload Me
ElseIf Option1.Value = True Then
       ysz = Val(Text1)
       If yscd = 1 Then
          ys1 = ysz
       ElseIf yscd = 2 Then
              ys2 = ysz
       ElseIf yscd = 3 Then
              ys3 = ysz
       ElseIf yscd = 4 Then
          ys4 = ysz
       End If
       cdgb = False
       Unload Me
End If
End Sub

Private Sub Command2_Click()
cdgb = True
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
List1.AddItem "黑色"
List1.AddItem "红色"
List1.AddItem "绿色"
List1.AddItem "黄色"
List1.AddItem "蓝色"
List1.AddItem "洋红色"
List1.AddItem "青色"
List1.AddItem "白色"
List1.AddItem "森林绿"
List1.AddItem "银色"
Text1.Enabled = False
Textr.Enabled = False
Textg.Enabled = False
Textb.Enabled = False
Option2.Value = True
Dim ran!(1 To 5)
Randomize
For i = 1 To 5
    ran(i) = Int(Rnd * (RGB(255, 255, 255) + 1))
Next i
Command1.BackColor = ran(1)
Command2.BackColor = ran(2)
Option1.BackColor = ran(3)
Option2.BackColor = ran(4)
Option3.BackColor = ran(5)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then cdgb = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   List1.Enabled = False
   Textr.Enabled = False
   Textg.Enabled = False
   Textb.Enabled = False
   Text1.Enabled = True
   Text1.SetFocus
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   List1.Enabled = True
   Textr.Enabled = False
   Textg.Enabled = False
   Textb.Enabled = False
   Text1.Enabled = False
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
   List1.Enabled = False
   Textr.Enabled = True
   Textg.Enabled = True
   Textb.Enabled = True
   Text1.Enabled = False
   Textr.SetFocus
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(Text1) > 16777215 Or Val(Text1) < 0 Then
   Text1 = ""
   Text1.SetFocus
End If
End Sub

Private Sub Textb_Keyup(KeyCode As Integer, Shift As Integer)
If (Val(Textb) > 255 Or Val(Textb) < 0) And Option3.Value = True Then
   Textb = ""
   Textb.SetFocus
End If
End Sub

Private Sub Textg_Keyup(KeyCode As Integer, Shift As Integer)
If (Val(Textg) > 255 Or Val(Textg) < 0) And Option3.Value = True Then
   Textg = ""
   Textg.SetFocus
End If
End Sub

Private Sub Textr_Keyup(KeyCode As Integer, Shift As Integer)
If (Val(Textr) > 255 Or Val(Textr) < 0) And Option3.Value = True Then
   Textr = ""
   Textr.SetFocus
End If
End Sub
