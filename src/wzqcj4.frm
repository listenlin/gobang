VERSION 5.00
Begin VB.Form Formdk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选取图片"
   ClientHeight    =   3630
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7335
   Icon            =   "wzqcj4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wzqcj4.frx":324A
   ScaleHeight     =   3630
   ScaleWidth      =   7335
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取  消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   3330
      Left            =   2880
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3030
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入保存文件名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   5760
      TabIndex        =   6
      Top             =   360
      Width           =   1440
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Formdk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Right(File1.Path, 1) <> "\" Then
   filename = File1.Path & "\" & File1.filename
Else
   filename = File1.Path & File1.filename
End If
If ave = 2 Or ave = 4 Then
   If Text1 = "" Then
      MsgBox "请输入文件名", 48, "提示"
      Text1.SetFocus
      Exit Sub
   End If
   If Right(filename, 1) <> "\" Then
      filename = filename & "\" & Text1
   Else
       filename = filename & Text1
   End If
End If
Unload Me
End Sub

Private Sub Command2_Click()
filename = ""
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If ave = 1 Then
Dim fi$
If Right(File1.Path, 1) <> "\" Then
   fi = File1.Path & "\" & File1.filename
Else
   fi = File1.Path & File1.filename
End If
Image1.Picture = LoadPicture(fi)
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
If ave = 1 Then
   Label1.Visible = False
   Text1.Visible = False
   Image1.Visible = True
ElseIf ave = 2 Then
   Label1.Visible = True
   Text1.Visible = True
   Image1.Visible = False
   File1.Visible = False
   File1.Pattern = ""
   Formdk.Caption = "保存棋谱"
ElseIf ave = 3 Then
       Label1.Visible = False
       Text1.Visible = False
       Image1.Visible = False
       File1.Pattern = "*.lsl"
       Formdk.Caption = "打开棋谱文件(后缀名lsl)"
ElseIf ave = 4 Then
       Label1.Visible = True
       Text1.Visible = True
       Image1.Visible = False
       File1.Visible = False
       File1.Pattern = ""
       Formdk.Caption = "保存文档"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
   filename = ""
End If
End Sub
