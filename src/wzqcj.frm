VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formzjm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "ɭ��������"
   ClientHeight    =   8610
   ClientLeft      =   1935
   ClientTop       =   1755
   ClientWidth     =   17235
   DrawWidth       =   3
   Icon            =   "wzqcj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   17235
   Begin MSComDlg.CommonDialog com 
      Left            =   14760
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox lisdh 
      Height          =   1335
      Left            =   13920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Timer Tmrxz 
      Interval        =   250
      Left            =   14040
      Top             =   7800
   End
   Begin VB.TextBox Textdh 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   31
      Top             =   5040
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Win 
      Index           =   1
      Left            =   2040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picsta 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   17205
      TabIndex        =   30
      Top             =   8310
      Width           =   17235
      Begin VB.PictureBox Picip 
         Height          =   255
         Left            =   2640
         ScaleHeight     =   195
         ScaleWidth      =   3315
         TabIndex        =   38
         Top             =   0
         Width           =   3375
         Begin VB.Label Labip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����ϸ��"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.PictureBox Picxz 
         Height          =   255
         Left            =   9720
         ScaleHeight     =   195
         ScaleWidth      =   1875
         TabIndex        =   36
         Top             =   0
         Width           =   1935
         Begin VB.Label Labelxz 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   37
            Top             =   0
            Width           =   1410
         End
      End
      Begin VB.Label Labelts 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6240
         TabIndex        =   35
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Labeldlm 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label Labelzb 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12840
         TabIndex        =   33
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Labelsj 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   14520
         TabIndex        =   32
         Top             =   0
         Width           =   1995
      End
   End
   Begin VB.TextBox Textip 
      Height          =   270
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Comlj 
      Caption         =   "�� �� �� �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Comjl 
      Caption         =   "�� �� �� �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Win 
      Index           =   0
      Left            =   2040
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   7920
      Width           =   11415
   End
   Begin VB.VScrollBar VS 
      Height          =   8175
      Left            =   13560
      TabIndex        =   24
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   7815
      Left            =   2160
      ScaleHeight     =   7755
      ScaleWidth      =   11355
      TabIndex        =   23
      Top             =   120
      Width           =   11415
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawWidth       =   2
         Height          =   7575
         Left            =   240
         ScaleHeight     =   7515
         ScaleWidth      =   10875
         TabIndex        =   26
         Top             =   240
         Width           =   10935
         Begin VB.Image Imab 
            Appearance      =   0  'Flat
            Height          =   1155
            Index           =   1
            Left            =   3240
            Picture         =   "wzqcj.frx":324A
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1155
         End
         Begin VB.Image Imah 
            Height          =   1155
            Index           =   1
            Left            =   1320
            Picture         =   "wzqcj.frx":4D83
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1155
         End
      End
   End
   Begin VB.PictureBox Picb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13920
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   22
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox Pich 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13920
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   21
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Comby 
      Caption         =   "����������ɫ&J"
      Height          =   375
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Comhy 
      Caption         =   "����������ɫ&H"
      Height          =   375
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Comks 
      Caption         =   "��ʼ����Ϸ&U"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton Combc 
      Caption         =   "�� �� �� ��&S"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Timer Timersj 
      Interval        =   1000
      Left            =   14040
      Top             =   7080
   End
   Begin VB.CommandButton Comhq 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��     ��&R"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "�������ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
      Begin VB.CommandButton Comzdy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Զ��屳��&B"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �� �� �� ˮ "
         ForeColor       =   &H00404080&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   5040
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   4
         Left            =   120
         Picture         =   "wzqcj.frx":68B1
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �� �� �� ��  "
         ForeColor       =   &H00008080&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   3
         Left            =   120
         Picture         =   "wzqcj.frx":282C7
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  һ ֻ �� ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   2
         Left            =   120
         Picture         =   "wzqcj.frx":20D503
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �� �� �� ��"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   885
         Index           =   1
         Left            =   120
         Picture         =   "wzqcj.frx":27DC86
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   0
         Left            =   120
         Picture         =   "wzqcj.frx":44A8AF
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �� ˮ С ��"
         ForeColor       =   &H00FF80FF&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Timer Timerb 
      Interval        =   1000
      Left            =   16200
      Top             =   1920
   End
   Begin VB.Timer Timerh 
      Interval        =   1000
      Left            =   16200
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "��Ϸ��ʼ����ѡ��"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Labelsjb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13920
      TabIndex        =   8
      Top             =   3000
      Width           =   2340
      WordWrap        =   -1  'True
   End
   Begin VB.Label Labelsjh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13920
      TabIndex        =   7
      Top             =   1200
      Width           =   2340
      WordWrap        =   -1  'True
   End
   Begin VB.Label Labelbsb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15120
      TabIndex        =   6
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Lal2 
      AutoSize        =   -1  'True
      Caption         =   "�׷���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   13800
      TabIndex        =   5
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Lal1 
      AutoSize        =   -1  'True
      Caption         =   "�ڷ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   13800
      TabIndex        =   4
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Labelbsh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15120
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
   Begin VB.Menu ck 
      Caption         =   " �� �� �� ����&C��"
      Begin VB.Menu ksxyx 
         Caption         =   "��  ʼ  ��  ��   Ϸ"
         Shortcut        =   ^G
      End
      Begin VB.Menu scqp 
         Caption         =   "���ϴ���Ϸ����"
         Shortcut        =   ^N
      End
      Begin VB.Menu bcqp 
         Caption         =   "��    ��    ��    ��"
         Shortcut        =   ^V
      End
      Begin VB.Menu ckbc 
         Caption         =   "��    ��    ��    ��"
         Shortcut        =   ^Q
      End
      Begin VB.Menu tcyx 
         Caption         =   "��    ��    ��    Ϸ"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu yxsz 
      Caption         =   " �� Ϸ �� �ã�&Z��"
      Begin VB.Menu jmsd 
         Caption         =   "�� �� �� ɫ"
         Begin VB.Menu sjtx 
            Caption         =   "���ͼ��"
         End
         Begin VB.Menu cs 
            Caption         =   "��      ɫ"
         End
      End
      Begin VB.Menu qpsz 
         Caption         =   "�� �� �� ��"
         Begin VB.Menu qp 
            Caption         =   "9��9"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu qp 
            Caption         =   "11��11"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu qp 
            Caption         =   "13��13"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu qp 
            Caption         =   "15��15(��׼����)"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu qp 
            Caption         =   "17��17"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu qp 
            Caption         =   "19��19"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu qp 
            Caption         =   "21��21"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu qp 
            Caption         =   "23��23"
            Checked         =   -1  'True
            Index           =   8
         End
         Begin VB.Menu qp 
            Caption         =   "25��25"
            Checked         =   -1  'True
            Index           =   9
         End
      End
      Begin VB.Menu ys 
         Caption         =   "�� �� �� ɫ"
         Begin VB.Menu hfys 
            Caption         =   "�ڷ���ɫ"
         End
         Begin VB.Menu bfys 
            Caption         =   "�׷���ɫ"
         End
      End
      Begin VB.Menu yxbj 
         Caption         =   "�� Ϸ �� ��"
         Begin VB.Menu tjtp 
            Caption         =   "���ͼƬ"
         End
         Begin VB.Menu sdys 
            Caption         =   "�趨��ɫ"
         End
      End
      Begin VB.Menu znsz 
         Caption         =   "����������"
         Begin VB.Menu jgx 
            Caption         =   "�� �� ��"
            Checked         =   -1  'True
         End
         Begin VB.Menu fsx 
            Caption         =   "�� �� ��"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu fzqz 
         Caption         =   "�� �� �� ��"
         Checked         =   -1  'True
         Shortcut        =   ^F
      End
      Begin VB.Menu ztl 
         Caption         =   "�� ʾ ״ ̬ ��"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
      Begin VB.Menu qxts 
         Caption         =   "�� �� �� ʾ"
         Checked         =   -1  'True
         Shortcut        =   ^R
      End
      Begin VB.Menu bcts 
         Caption         =   "�� �� �� ʾ"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu dlwj 
      Caption         =   " �� ¼ �� �ң�&W��"
      Begin VB.Menu xgmm 
         Caption         =   "�� �� �� ��"
         Shortcut        =   ^M
      End
      Begin VB.Menu dzsj 
         Caption         =   "�� ս �� �� "
         Shortcut        =   ^D
      End
      Begin VB.Menu tcdl 
         Caption         =   "�� �� �� ¼"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu dzms 
      Caption         =   " �� ս ģ ʽ��&D��"
      Begin VB.Menu drdz 
         Caption         =   "�� �� �� ս"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu rjdz 
         Caption         =   "�� �� �� ս"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu wldz 
         Caption         =   "�� �� �� �� ս"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu yxxz 
      Caption         =   " �� Ϸ �� �ƣ�&A��"
      Begin VB.Menu sjxz 
         Caption         =   " ʱ �� �� ��"
         Begin VB.Menu shi 
            Caption         =   "2��00"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu shi 
            Caption         =   "5��00"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu shi 
            Caption         =   "7��00"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu shi 
            Caption         =   "10��00"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu shi 
            Caption         =   "15��00"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu shi 
            Caption         =   "20��00"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu shi 
            Caption         =   "�Զ���ʱ��"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu shi 
            Caption         =   "ȡ �� �� ��"
            Index           =   8
         End
      End
      Begin VB.Menu bsxz 
         Caption         =   " �� �� �� ��"
         Begin VB.Menu bu 
            Caption         =   "��40��"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu bu 
            Caption         =   "��60��"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu bu 
            Caption         =   "��80��"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu bu 
            Caption         =   "��100��"
            Checked         =   -1  'True
            Index           =   4
         End
      End
      Begin VB.Menu jsxz 
         Caption         =   " �� �� �� ��"
         Begin VB.Menu ssjs 
            Caption         =   " �� �� �� ��"
            Checked         =   -1  'True
         End
         Begin VB.Menu sijs 
            Caption         =   " �� �� �� ��"
            Checked         =   -1  'True
         End
         Begin VB.Menu cljs 
            Caption         =   " �� �� �� ��"
            Checked         =   -1  'True
         End
         Begin VB.Menu qxjs 
            Caption         =   " ȡ �� �� ��"
         End
      End
   End
   Begin VB.Menu zykj 
      Caption         =   " ְ ҵ �� �֣�&K��"
      Begin VB.Menu zzdf 
         Caption         =   "ֱָ�򷨿���"
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   1
         End
         Begin VB.Menu zz 
            Caption         =   " Ϫ  ��  ��"
            Index           =   2
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   3
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   4
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   5
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   6
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   7
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   8
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   9
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   10
         End
         Begin VB.Menu zz 
            Caption         =   " ��  ��  ��"
            Index           =   11
         End
         Begin VB.Menu zz 
            Caption         =   " ɽ  ��  ��"
            Index           =   12
         End
      End
      Begin VB.Menu xzdf 
         Caption         =   "бָ�򷨿���"
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   1
         End
         Begin VB.Menu xz 
            Caption         =   " Ͽ  ��  ��"
            Index           =   2
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   3
         End
         Begin VB.Menu xz 
            Caption         =   " ˮ  ��  ��"
            Index           =   4
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   5
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   6
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   7
         End
         Begin VB.Menu xz 
            Caption         =   " �  ��  �� "
            Index           =   8
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   9
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  �� "
            Index           =   10
         End
         Begin VB.Menu xz 
            Caption         =   " б  ��  ��"
            Index           =   11
         End
         Begin VB.Menu xz 
            Caption         =   " ��  ��  ��"
            Index           =   12
         End
      End
   End
   Begin VB.Menu gyyx 
      Caption         =   " �� Ϸ �� ����&H��"
      Begin VB.Menu wzqjj 
         Caption         =   "�� �� �� �� ��"
         Shortcut        =   ^W
      End
      Begin VB.Menu yxsm 
         Caption         =   "��  Ϸ  ˵  ��"
         Shortcut        =   ^E
      End
      Begin VB.Menu gywzq 
         Caption         =   "�� �� ɭ �� �� �� ��"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu tc1 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu tcys 
         Caption         =   "��  ��  ��  ɫ"
         Begin VB.Menu hy 
            Caption         =   "�ڷ���ɫ"
         End
         Begin VB.Menu bye 
            Caption         =   "�׷���ɫ"
         End
      End
      Begin VB.Menu chakan 
         Caption         =   "��  ��  ��  ��"
      End
      Begin VB.Menu szqp 
         Caption         =   "��  ��  ��  ��"
         Begin VB.Menu pq 
            Caption         =   "9��9"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu pq 
            Caption         =   "11��11"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu pq 
            Caption         =   "13��13"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu pq 
            Caption         =   "15��15(��׼����)"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu pq 
            Caption         =   "17��17"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu pq 
            Caption         =   "19��19"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu pq 
            Caption         =   "21��21"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu pq 
            Caption         =   "23��23"
            Checked         =   -1  'True
            Index           =   8
         End
         Begin VB.Menu pq 
            Caption         =   "25��25"
            Checked         =   -1  'True
            Index           =   9
         End
      End
      Begin VB.Menu youxibeijing 
         Caption         =   "��  Ϸ  ��  ��"
         Begin VB.Menu tupian 
            Caption         =   "���ͼƬ"
         End
         Begin VB.Menu yanse 
            Caption         =   "�趨��ɫ"
         End
      End
      Begin VB.Menu tchq 
         Caption         =   "��            ��"
      End
      Begin VB.Menu tcbc 
         Caption         =   "��  ��  ��  ��"
      End
      Begin VB.Menu tcks 
         Caption         =   "��  ʼ  ��  ��  Ϸ"
      End
      Begin VB.Menu qxxz 
         Caption         =   "ȡ �� �� �� �� ��"
      End
      Begin VB.Menu xztp 
         Caption         =   "ж �� �� �� ͼ Ƭ"
      End
   End
   Begin VB.Menu tanchu 
      Caption         =   "����1"
      Visible         =   0   'False
      Begin VB.Menu xdtp 
         Caption         =   "ѡ����ͼƬΪ����"
      End
      Begin VB.Menu ggtp 
         Caption         =   "�� �� �� ͼ Ƭ"
      End
   End
End
Attribute VB_Name = "Formzjm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wz$(24, 24), wez$(24, 24)          '��������λ�����꼰��������
Dim hzb%, zzb%                         '�����꣬������
Dim lzd$                               '���ӵ��ַ���¼
Dim qxh$, qxb$                         '���������ַ���¼
Dim dhz%                               '���������ߺ�����
Dim hz$(1 To 313), bz$(1 To 313)       '��¼�ڼ�����������Ϣ��hzΪ�ڷ�����ԣ�bzΪ�׷������
Dim slz As Boolean                     '�жϸ��ķ����ӣ�tureΪ���ӣ�falseΪ����(˭����)
Dim bsb%                               '�׷�����Ĳ���
Dim bsh%                               '�ڷ�����Ĳ���
Dim sjb%                               '�׷������ʱ��
Dim sjh%                               '�ڷ������ʱ��
Dim ys1 As Long                        '�������ɫֵ
Dim ys2 As Long                        '�������ɫֵ
Dim tsbc As Boolean, bt As Boolean
Dim xzs%, xzb%                         'ʱ�䲽�������ƴ�С
Dim xzjs$(1 To 5)    '���Ƶ���ʾ����
Dim ma!, mb!, mc!, mdd!, pt1!, pt2!, v1!, h1!   'ͼƬ������ر���
Dim windowsh!, windowsw!, res As Boolean
Dim zhc$, laizou As Boolean, inde%
Dim threej$, fourj$
Dim thtw As Boolean, tto As Boolean, gyg As Boolean, oot As Boolean
Dim rrz As Boolean, ggz As Boolean, ssz As Boolean
Private Sub bcqp_Click()
Call Combc_Click
End Sub

Private Sub bcts_Click()
If bcts.Checked = True Then
   bcts.Checked = False
   bt = False
ElseIf bcts.Checked = False Then
       bcts.Checked = True
       bt = True
End If
End Sub

Private Sub bu_Click(index As Integer)
For i = 1 To 16
    If bu(i).Visible = True Then
       If i = index And i <= 14 Then
          bu(i).Checked = True
          bu(i).Enabled = False
          xzb = Val(Mid(bu(i).Caption, 2, 3))
          xzjs(2) = "��" & xzb & "��"
          If md = 3 Then
             If laizou = False Then
                If Win(1).State = sckConnected Then
                   If index >= 10 Then
                      Win(1).SendData (xzb & index & "xb")
                   Else
                       Win(1).SendData (xzb & "9" & index & "xb")
                   End If
                End If
             End If
          End If
       Else
           bu(i).Checked = False
           bu(i).Enabled = True
       End If
    End If
Next i
If index = 15 Then
   If laizou = False Then
      Dim b!
      b = Val(InputBox("���������Ʋ�����", "���Ʋ�������"))
      If b > 0 And b < 32767 Then
         xzb = Round(b)
         bu(15).Checked = True
         bu(15).Enabled = True
         xzjs(2) = "��" & xzb & "��"
         If md = 3 Then
            If laizou = False Then
               If Win(1).State = sckConnected Then
                  Win(1).SendData (xzb & index & "xb")
               End If
            End If
         End If
      End If
   ElseIf laizou = True Then
          bu(15).Checked = True
          bu(15).Enabled = True
   End If
End If
If index = 16 Then
   bu(16).Checked = True
   bu(16).Enabled = False
   xzb = 0
   xzjs(2) = "�����Ʋ���"
   If md = 3 Then
      If laizou = False Then
         If Win(1).State = sckConnected Then
            Win(1).SendData (0 & index & "xb")
         End If
      End If
   End If
End If
End Sub

Private Sub chakan_Click()
Call ckbc_Click
End Sub
Private Sub ckbc_Click()   '�򿪱��������
If bsh >= 1 Or bsb >= 1 Then
   If bt = True Then
      If tsbc = False Then
         Dim ad%
         ad = MsgBox("�Ƿ񱣴����ף�", 36, "��ʾ")
         If ad = vbYes Then
            Call Combc_Click
         End If
      End If
   End If
End If
Picture1.Enabled = True
If md <> 3 Or (md = 3 And zhc = "zhu") Then
   com.CancelError = True
   On Error GoTo errhandler
   com.Filter = "*.lsl"
   com.ShowOpen
   If Right(com.FileName, 3) <> "lsl" Then
      MsgBox "����غ�׺Ϊ��lsl��������", 48, "��ʾ"
      Exit Sub
   End If
   If bsh >= 10 Or bsb >= 10 Then
      bsh = 1: bsb = 0
      Call Comks_Click(1)
   Else
       bsh = 0: bsb = 0
       Call Comks_Click(0)
   End If
   tsbc = True
   ys.Enabled = False
   Comhy.Visible = False
   Comby.Visible = False
   tcys.Enabled = False
   qpsz.Enabled = False
   szqp.Enabled = False
   zykj.Enabled = False
   Dim cv As save, jlh%, h%, z%, send$, bshb$, jr$, yj$
   Open com.FileName For Random As #2 Len = Len(cv)
        jlh = 1
        Do Until LOF(2) / Len(cv) < jlh
           Get #2, jlh, cv
           If jlh = 1 Then
              sjh = cv.sjh: sjb = cv.sjb
              ys1 = cv.ysh: ys2 = cv.ysb
              Pich.Scale (0, Pich.Height)-(Pich.Width, 0)
              For i = 0 To Pich.Height / 2
                  Pich.Circle (Pich.Width / 2, Pich.Width / 2), i, ys1
              Next i
              Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
              For j = 0 To Picb.Height / 2
                  Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, ys2
              Next j
              If md = 1 Then
                 Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
                 If sjh = 0 Then
                    sjh = sjb
                    Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                 Else
                     Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                 End If
              ElseIf md = 2 Then
                     Labelsjb.Caption = dl.mz & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
              ElseIf md = 3 Then
                     Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
                     If sjh = 0 Then
                        sjh = sjb
                        Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                     Else
                         Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                     End If
                     If Win(1).State = sckConnected Then
                        send = "," & sjh & "," & sjb & "," & ys1 & "," & ys2
                     End If
              End If
           End If
           Select Case Trim(cv.zbh)
                   Case ""
                        If cv.zbb = "1234" Then
                           bsh = jlh - 1: bsb = jlh - 1
                           Labelbsb.Caption = "��" & bsb & "��"
                           Labelbsh.Caption = "��" & bsh & "��"
                           If Win(1).State = sckConnected Then
                              bshb = "," & bsh & "," & bsb
                           End If
                           If md <> 2 Then
                              Timerh.Enabled = True
                           End If
                           slz = True
                           Option1.Value = True
                           Frame1.Caption = "��Ϸ�У�����ѡ��"
                           Frame1.Enabled = False
                           Call jstr(Trim(Str(cv.sjb)), h, z)
                           If z > dhz Then
                              If Win(1).State = sckConnected Then
                                 jr = z
                              End If
                              MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                              bsh = 0: bsb = 0
                              Call Comks_Click(0)
                              slz = True
                              Picture2.Enabled = True
                              Exit Do
                           Else
                               jr = "0"
                           End If
                           If h = 1 Then
                              Timerh.Enabled = False
                              Picture2.Enabled = False
                              If Win(1).State = sckConnected Then
                                 yj = "b"
                                 MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                              Else
                                  MsgBox "������" & " " & "�׷�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                              End If
                           ElseIf h = 2 Then
                                  Timerh.Enabled = False
                                  Picture2.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & "���" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                           ElseIf h = 3 Then
                                  Timerh.Enabled = False
                                  Picture2.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                           Else
                               If md = 2 Then
                                  Dim ah%, az%
                                  Call jstr(autolz(wz()), ah, az)
                                  Call hqz(ah, az, bsh + 1, slz)
                                  wz(ah, az) = "����"
                                  bsh = bsh + 1
                                  hz(bsh) = bstr(ah, az)
                                  Labelbsh.Caption = "��" & bsh & "��"
                                  If fzqz.Checked = False Then
                                     Call fk(slz)
                                  End If
                                  slz = False
                                  Timerb.Enabled = True
                                  Labelsjh.Caption = ""
                                  If pdwz(wz(), ah, az, "����") = True Then
                                  End If
                               ElseIf md = 3 Then
                                      Timerh.Enabled = True
                                      If Win(1).State = sckConnected Then
                                         yj = "hs"
                                         MsgBox "�����״˲��������ӣ������ӣ�", 48, "��ʾ"
                                      End If
                               End If
                           End If
                        ElseIf Trim(cv.zbb) <> "" Then
                               bsh = jlh - 1: bsb = jlh
                               Labelbsb.Caption = "��" & bsb & "��"
                               Labelbsh.Caption = "��" & bsh & "��"
                               If Win(1).State = sckConnected Then
                                  bshb = "," & bsh & "," & bsb
                               End If
                               If md <> 2 Then
                                  Timerh.Enabled = True
                               End If
                               slz = True
                               Option2.Value = True
                               Frame1.Caption = "��Ϸ�У�����ѡ��"
                               Frame1.Enabled = False
                               bz(jlh) = cv.zbb
                               Call jstr(cv.zbb, h, z)
                               wz(h, z) = "����"
                               Call hqz(h, z, bsb, Not slz)
                               If Win(1).State = sckConnected Then
                                  send = send + "," & cv.zbb & jlh & "h"
                               End If
                               Call jstr(Trim(Str(cv.sjb)), h, z)
                               If z > dhz Then
                                  If Win(1).State = sckConnected Then
                                     jr = z
                                  End If
                                  MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                                  bsh = 0: bsb = 0
                                  Call Comks_Click(0)
                                  slz = True
                                  Picture2.Enabled = True
                                  Exit Do
                               Else
                                   jr = "0"
                               End If
                               If h = 1 Then
                                  Picture2.Enabled = False
                                  Timerh.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & "�׷�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                               ElseIf h = 2 Then
                                      Picture2.Enabled = False
                                      Timerh.Enabled = False
                                      If Win(1).State = sckConnected Then
                                         yj = "b"
                                         MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                      Else
                                          MsgBox "������" & " " & "���" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                      End If
                               ElseIf h = 3 Then
                                      Picture2.Enabled = False
                                      Timerh.Enabled = False
                                      If Win(1).State = sckConnected Then
                                         yj = "b"
                                         MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                      Else
                                          MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                      End If
                               Else
                                   If md = 2 Then
                                      Call jstr(autolz(wz()), ah, az)
                                      Call hqz(ah, az, bsh + 1, slz)
                                      wz(ah, az) = "����"
                                      bsh = bsh + 1
                                      hz(bsh) = bstr(ah, az)
                                      Labelbsh.Caption = "��" & bsh & "��"
                                      If fzqz.Checked = False Then
                                         Call fk(slz)
                                      End If
                                      slz = False
                                      Timerb.Enabled = True
                                      Labelsjh.Caption = ""
                                      If pdwz(wz(), ah, az, "����") = True Then
                                      End If
                                   ElseIf md = 3 Then
                                          Timerh.Enabled = True
                                          If Win(1).State = sckConnected Then
                                             yj = "hs"
                                             MsgBox "�����״˲��������ӣ������ӣ�", 48, "��ʾ"
                                          End If
                                   End If
                               End If
                        End If
                   Case "1234"
                        bsb = jlh - 1: bsh = jlh - 1
                        Labelbsb.Caption = "��" & bsb & "��"
                        Labelbsh.Caption = "��" & bsh & "��"
                        If Win(1).State = sckConnected Then
                           bshb = "," & bsh & "," & bsb
                        End If
                        Timerb.Enabled = True
                        slz = False
                        Option2.Value = True
                        Frame1.Caption = "��Ϸ�У�����ѡ��"
                        Frame1.Enabled = False
                        Call jstr(Trim(Str(cv.sjh)), h, z)
                        If z > dhz Then
                           If Win(1).State = sckConnected Then
                              jr = z
                           End If
                           MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                           bsh = 0: bsb = 0
                           Call Comks_Click(0)
                           slz = True
                           Picture2.Enabled = True
                           Exit Do
                        Else
                            jr = "0"
                        End If
                        If h = 1 Then
                           Timerb.Enabled = False
                           Picture2.Enabled = False
                           If Win(1).State = sckConnected Then
                              yj = "h"
                              MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                           Else
                               MsgBox "������" & " " & "�ڷ�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                           End If
                        ElseIf h = 2 Then
                               Timerb.Enabled = False
                               Picture2.Enabled = False
                               If Win(1).State = sckConnected Then
                                  yj = "h"
                                  MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                               Else
                                  MsgBox "������" & " " & "����" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                               End If
                        ElseIf h = 3 Then
                               Timerb.Enabled = False
                               Picture2.Enabled = False
                               If Win(1).State = sckConnected Then
                                  yj = "h"
                                  MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                               Else
                                   MsgBox "������" & " " & Lal1 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                               End If
                        ElseIf md = 3 Then
                               Timerb.Enabled = True
                               If Win(1).State = sckConnected Then
                                  yj = "bs"
                                  MsgBox "�����״˲��ɶԷ����ӣ���ȴ���", 48, "��ʾ"
                               End If
                        End If
                   Case Else
                       If Trim(cv.zbb) = "" Then
                          bsh = jlh: bsb = jlh - 1
                          Labelbsb.Caption = "��" & bsb & "��"
                          Labelbsh.Caption = "��" & bsh & "��"
                          If Win(1).State = sckConnected Then
                             bshb = "," & bsh & "," & bsb
                          End If
                          Timerb.Enabled = True
                          slz = False
                          Option1.Value = True
                          Frame1.Caption = "��Ϸ�У�����ѡ��"
                          Frame1.Enabled = False
                          hz(jlh) = cv.zbh
                          Call jstr(cv.zbh, h, z)
                          wz(h, z) = "����"
                          Call hqz(h, z, bsh, Not slz)
                          If Win(1).State = sckConnected Then
                             send = send + "," & cv.zbh & jlh & "b"
                          End If
                          Call jstr(Trim(Str(cv.sjh)), h, z)
                          If z > dhz Then
                              If Win(1).State = sckConnected Then
                                 jr = z
                              End If
                              MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                              bsh = 0: bsb = 0
                              Call Comks_Click(0)
                              Picture2.Enabled = True
                              slz = True
                              Exit Do
                          Else
                              jr = "0"
                          End If
                          If h = 1 Then
                             Timerb.Enabled = False
                             Picture2.Enabled = False
                             If Win(1).State = sckConnected Then
                                yj = "h"
                                MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                             Else
                                 MsgBox "������" & " " & "�ڷ�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                             End If
                          ElseIf h = 2 Then
                                 Timerb.Enabled = False
                                 Picture2.Enabled = False
                                 If Win(1).State = sckConnected Then
                                    yj = "h"
                                    MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 Else
                                     MsgBox "������" & " " & "����" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 End If
                          ElseIf h = 3 Then
                                 Timerb.Enabled = False
                                 Picture2.Enabled = False
                                 If Win(1).State = sckConnected Then
                                    yj = "h"
                                    MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 Else
                                    MsgBox "������" & " " & Lal1 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 End If
                          ElseIf md = 3 Then
                                 Timerb.Enabled = True
                                 If Win(1).State = sckConnected Then
                                    yj = "bs"
                                    MsgBox "�����״˲��ɶԷ����ӣ���ȴ���", 48, "��ʾ"
                                 End If
                          End If
                       Else
                           hz(jlh) = cv.zbh
                           Call jstr(cv.zbh, h, z)
                           wz(h, z) = "����"
                           Call hqz(h, z, jlh, tsbc)
                           If Win(1).State = sckConnected Then
                              send = send + "," & cv.zbh & jlh & "b"
                           End If
                           '///////////////////////////////////
                           bz(jlh) = cv.zbb
                           Call jstr(cv.zbb, h, z)
                           wz(h, z) = "����"
                           Call hqz(h, z, jlh, Not tsbc)
                           If Win(1).State = sckConnected Then
                              send = send + "," & cv.zbb & jlh & "h"
                           End If
                       End If
           End Select
           jlh = jlh + 1
        Loop
   Close #2
   If Win(1).State = sckConnected Then
      send = jr & bshb & "," & yj & send & "lp"
      Win(1).SendData send
   End If
errhandler:
End If
End Sub

Private Sub cljs_Click()
cljs.Checked = Not cljs.Checked
If cljs.Checked = True Then
   xzjs(5) = "�г�������"
Else
    xzjs(5) = ""
End If
If md = 3 And laizou = False Then
   If Win(1).State = sckConnected And zhc = "bei" Then
      Win(1).SendData ("cljs" & "js")
   End If
End If
End Sub

Private Sub Combc_Click()       '�������
If bsh = 0 And bsb = 0 Then
   MsgBox "�����ױ���", 48, "��ʾ"
   Exit Sub
End If
If Timerh = True Then Timerh = False
If Timerb = True Then Timerb = False
com.CancelError = True
On Error GoTo errhandler
com.ShowSave
If com.FileName <> "" Then
       Dim cu As save, jlh%, sy%
       If tsbc = True Then
          sy = md
       ElseIf tsbc = False Then
              sy = 0
       End If
       Open com.FileName & ".lsl" For Random As #1 Len = Len(cu)
           jlh = 1
           Do Until hz(jlh) = "" And bz(jlh) = ""
              cu.zbh = hz(jlh): cu.ysh = ys1
              cu.zbb = bz(jlh): cu.ysb = ys2
              cu.sjh = sjh: cu.sjb = sjb
              If hz(jlh) = "" Then
                    cu.sjb = Val(bstr(sy, dhz))
              End If
              If bz(jlh) = "" Then
                 cu.sjh = Val(bstr(sy, dhz))
              End If
              Put #1, jlh, cu
              jlh = jlh + 1
           Loop
           If hz(jlh - 1) <> "" And bz(jlh - 1) <> "" Then
              If slz = True Then
                 cu.zbh = "": cu.ysh = 0
                 cu.zbb = "1234": cu.ysb = 0
                 cu.sjh = 0: cu.sjb = Val(bstr(sy, dhz))
                 Put #1, jlh, cu
              ElseIf slz = False Then
                     cu.zbh = "1234": cu.ysh = 0
                     cu.zbb = "": cu.ysb = 0
                     cu.sjh = Val(bstr(sy, dhz)): cu.sjb = 0
                     Put #1, jlh, cu
              End If
           End If
       Close #1
       MsgBox "�����ѱ�����" & com.FileName & ".lsl", 48, "��ʾ"
       tsbc = True
End If
errhandler:
   If slz = True Then
      Timerh.Enabled = True
   Else
       Timerb.Enabled = True
   End If
End Sub

Private Sub comzdy_Click()
Call tjtp_Click
End Sub


Private Sub cs_Click()
com.CancelError = True
On Error GoTo errhandler
com.ShowColor
    Formzjm.BackColor = com.Color
    Frame1.BackColor = com.Color
    Frame2.BackColor = com.Color
    Option1.BackColor = com.Color
    Option2.BackColor = com.Color
    Dim an!
    Randomize
    an = Int(Rnd * RGB(255, 255, 255) + 1)
    Comzdy.BackColor = an
    Comhy.BackColor = an
    Comby.BackColor = an
    Comhq.BackColor = an
    Combc.BackColor = an
    Comks(1).BackColor = an
    Comjl.BackColor = an
    Comlj.BackColor = an
    Pich.Scale (0, 10)-(10, 0)
    Pich.Cls
    Pich.BackColor = Me.BackColor
    For i = 0 To 100
        Pich.Circle (5, 5), i / 20, ys1
    Next i
    Picb.Scale (0, 10)-(10, 0)
    Picb.Cls
    Picb.BackColor = Me.BackColor
    For i = 0 To 100
        Picb.Circle (5, 5), i / 20, ys2
    Next i
errhandler:
End Sub

Private Sub dzsj_Click()
Formsj.Show 1
End Sub

Private Sub Form_Activate()
Call wzqsm
End Sub

Private Sub Form_Resize()
On Error Resume Next
If res = True Then
   Me.Width = windowsw
   Me.Height = windowsh
End If
End Sub

Private Sub fsx_Click()
fsx.Checked = True: fsx.Enabled = False
jgx.Checked = False: jgx.Enabled = True
End Sub

Private Sub fzqz_Click()
fzqz.Checked = Not fzqz.Checked
Dim he%, ze%, i%
If fzqz.Checked = True Then
   ys.Enabled = False
   tcys.Enabled = False
   Comhy.Visible = False
   Comby.Visible = False
   Call xztp_Click
   Picture1.BackColor = RGB(255, 255, 255)
   Picture1.Cls
   Call hqp
   Pich.Scale (0, 10)-(10, 0)
   Pich.Cls
   Pich.BackColor = Me.BackColor
   For i = 0 To 100
       Pich.Circle (5, 5), i / 20, RGB(0, 0, 0)
   Next i
   Picb.Scale (0, 10)-(10, 0)
   Picb.Cls
   Picb.BackColor = Me.BackColor
   For i = 0 To 100
       Picb.Circle (5, 5), i / 20, RGB(255, 255, 255)
   Next i
Else
    For i = 1 To 313
        Imah(i).Visible = False
        Imab(i).Visible = False
    Next i
    Picture1.BackColor = &H80FFFF
    Call hqp
    If (bsh = 0 And bsb = 0) And md <> 3 Then
       ys.Enabled = True
       tcys.Enabled = True
       Comhy.Visible = True
       Comby.Visible = True
    End If
    Pich.Scale (0, 10)-(10, 0)
    Pich.Cls
    Pich.BackColor = Me.BackColor
    For i = 0 To 100
        Pich.Circle (5, 5), i / 20, ys1
    Next i
    Picb.Scale (0, 10)-(10, 0)
    Picb.Cls
    Picb.BackColor = Me.BackColor
    For i = 0 To 100
        Picb.Circle (5, 5), i / 20, ys2
    Next i
End If
If bsh >= 1 Then
   For i = 1 To bsh
       Call jstr(hz(i), he, ze)
       Call hqz(he, ze, i, True)
   Next i
End If
If bsb >= 1 Then
   For i = 1 To bsb
       Call jstr(bz(i), he, ze)
       Call hqz(he, ze, i, False)
   Next i
End If
End Sub

Private Sub ggtp_Click()
com.CancelError = True
On Error GoTo errhandler
com.ShowOpen
com.Filter = "*.jpg;*.bmp;*.gif"
If Right(com.FileName, 3) <> "JPG" And Right(com.FileName, 3) <> "BMP" And Right(com.FileName, 3) <> "GIF" And Right(com.FileName, 3) <> "jpg" And Right(com.FileName, 3) <> "bmp" And Right(com.FileName, 3) <> "gif" Then
   MsgBox "����غ�׺��Ϊ��jpg,bmp��gif����ͼƬ", 48, "��ʾ"
   Exit Sub
End If
Image1(inde) = LoadPicture(com.FileName)
Dim hx$
For i = Len(com.FileName) To 1 Step -1
    hx = Mid(com.FileName, i, 1)
    If hx = "\" Then
       Exit For
   End If
Next i
hx = Right(com.FileName, Len(com.FileName) - i)
hx = Left(hx, Len(hx) - 4)
Label1(inde) = hx
errhandler:
End Sub

Private Sub gywzq_Click()
frmAbout.Show 1
End Sub

Private Sub Imab_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   If Picture1 = LoadPicture("") Then
      xztp.Enabled = False
   Else
       xztp.Enabled = True
   End If
   PopupMenu tc1, 0
End If
End Sub

Private Sub Imab_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hc%, zc%
Call jstr(bz(index), hc, zc)
Labelzb.Caption = Chr(65 + hc) & zc + 1 & " " & lzd
End Sub

Private Sub Image1_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   inde = index
   PopupMenu tanchu, 0
End If
End Sub

Private Sub Imah_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   If Picture1 = LoadPicture("") Then
      xztp.Enabled = False
   Else
       xztp.Enabled = True
   End If
   PopupMenu tc1, 0
End If
End Sub

Private Sub Imah_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hc%, zc%
Call jstr(hz(index), hc, zc)
Labelzb.Caption = Chr(65 + hc) & zc + 1 & " " & lzd
End Sub

Private Sub jgx_Click()
jgx.Checked = True: jgx.Enabled = False
fsx.Checked = False: fsx.Enabled = True
End Sub

Private Sub ksxyx_Click()
If bsh >= 10 Or bsb >= 10 Then
   Call Comks_Click(1)
Else
    Call Comks_Click(0)
End If
End Sub

Private Sub lisdh_Change()
lisdh.SelStart = Len(lisdh)
End Sub

Private Sub lisdh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   com.CancelError = True
   On Error GoTo errhandler
   com.Flags = cdlCFEffects Or cdlCFBoth
   com.ShowFont
   lisdh.Font.Name = com.FontName
   lisdh.Font.Size = com.FontSize
   lisdh.Font.Bold = com.FontBold
   lisdh.Font.Italic = com.FontItalic
   lisdh.Font.Underline = com.FontUnderline
   lisdh.FontStrikethru = com.FontStrikethru
   lisdh.ForeColor = com.Color
errhandler:
End If
End Sub


Private Sub Option1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If md = 2 Then
   Option1.Value = True
   Call hqz((dhz - 1) / 2, (dhz - 1) / 2, 1, True)
   wz((dhz - 1) / 2, (dhz - 1) / 2) = "����"
   bsh = 1
   hz(bsh) = bstr((dhz - 1) / 2, (dhz - 1) / 2)
   Labelbsh.Caption = "��" & bsh & "��"
   Timerb.Enabled = True
   slz = False
   Frame1.Caption = "��Ϸ�У�����ѡ��"
   Frame1.Enabled = False
   Comhy.Visible = False
   Comby.Visible = False
   ys.Enabled = False
   tcys.Enabled = False
   dzms.Enabled = False
   qpsz.Enabled = False
   szqp.Enabled = False
   zykj.Enabled = False
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hc%, zc%
Call pd(X, Y, hc, zc)
Labelzb.Caption = Chr(65 + hc) & zc + 1 & " " & lzd
End Sub

Private Sub pq_Click(index As Integer)
Call qp_Click(index)
End Sub

Private Sub qp_Click(index As Integer)
For i = 1 To 9
    If i = index Then
       qp(i).Checked = True
       qp(i).Enabled = False
       dhz = Val(Left(qp(i).Caption, 2))
       pt1 = dhz * 10 / Picture2.Height
       pt2 = dhz * 10 / Picture2.Width
       Me.Caption = "ɭ��������" & "��" & qp(i).Caption & "��"
       pq(i).Checked = True
       pq(i).Enabled = False
       If md = 3 Then
          If zhc = "zhu" Then
             If Win(1).State = sckConnected Then
                Win(1).SendData (index & "qp")
             End If
          End If
       End If
       For j = 40 To 25 ^ 2 / 2 Step 20
           If j < dhz ^ 2 / 2 Then
              bu((j - 40) / 20 + 1).Visible = True
              bu((j - 40) / 20 + 1).Caption = "��" & j & "��"
              bu((j - 40) / 20 + 1).Checked = False
              bu((j - 40) / 20 + 1).Enabled = True
           Else
               bu((j - 40) / 20 + 1).Visible = False
           End If
       Next j
    Else
        pq(i).Checked = False
        pq(i).Enabled = True
        qp(i).Checked = False
        qp(i).Enabled = True
    End If
Next i
If Picture1 <> LoadPicture("") Then
   Call chqp
Else
    Picture1.Height = Picture2.Height
    Picture1.Width = Picture2.Width
    Picture1.Scale (0, dhz * 10)-(dhz * 10, 0)
End If
Call hqp
End Sub


Private Sub qxjs_Click()
ssjs.Checked = False: threej = "": xzjs(3) = ""
sijs.Checked = False: fourj = "": xzjs(4) = ""
cljs.Checked = False: xzjs(5) = ""
If md = 3 Then
   If Win(1).State = sckConnected And laizou = False Then
      Win(1).SendData ("qxjs" & "qx")
   End If
End If
End Sub

Private Sub qxts_Click()
qxts.Checked = Not qxts.Checked
End Sub

Private Sub qxxz_Click()
Call bu_Click(16)
Call shi_Click(8)
If md = 1 Or md = 2 Then
   Call qxjs_Click
End If
If md = 3 Then
   If zhc = "bei" Then
      Call qxjs_Click
   End If
End If
If md = 3 Then
   If Win(1).State = sckConnected Then
      Win(1).SendData ("qx")
   End If
End If
End Sub

Private Sub shi_Click(index As Integer)
For i = 1 To 8
    If i = index And i <= 6 Then
       xzs = Val(Left(shi(index).Caption, 2))
       xzs = xzs * 60
       shi(i).Checked = True
       shi(i).Enabled = False
       xzjs(1) = "��ʱ" & xzs & "��"
       If md = 3 Then
          If laizou = False Then
             If Win(1).State = sckConnected Then
                Win(1).SendData (xzs & index & "xs")
             End If
          End If
       End If
    Else
        shi(i).Checked = False
        shi(i).Enabled = True
    End If
Next i
If index = 7 Then
   If laizou = False Then
      Dim a!
      a = Val(InputBox("���������������ΪС��", "�Զ���ʱ������"))
      If a > 0 And a < 32767 Then
         xzs = Round(a * 60)
         shi(7).Checked = True
         shi(7).Enabled = True
         xzjs(1) = "��ʱ" & xzs & "��"
         If md = 3 Then
            If laizou = False Then
               If Win(1).State = sckConnected Then
                  Win(1).SendData (xzs & index & "xs")
               End If
            End If
         End If
      End If
   ElseIf laizou = True Then
          shi(7).Checked = True
          shi(7).Enabled = True
   End If
ElseIf index = 8 Then
       xzs = 0
       For i = 1 To 7
           shi(i).Checked = False
           shi(i).Enabled = True
       Next i
       xzjs(1) = "������ʱ��"
       shi(8).Checked = True
       shi(8).Enabled = False
       If md = 3 Then
          If laizou = False Then
             If Win(1).State = sckConnected Then
                Win(1).SendData (0 & index & "xs")
             End If
          End If
       End If
End If
End Sub


Private Sub sijs_Click()
sijs.Checked = Not sijs.Checked
If sijs.Checked = False Then
   fourj = ""
   xzjs(4) = ""
Else
    Call qxzt
    xzjs(4) = "�����Ľ���"
End If
If md = 3 And laizou = False Then
   If Win(1).State = sckConnected And zhc = "bei" Then
      Win(1).SendData ("sijs" & "js")
   End If
End If
End Sub

Private Sub sjtx_Click()
Dim nz%, r!, g!, b!
Randomize
nz = Int(Rnd * 4)
If nz = 0 Then
       For j = 1 To Me.Height
           r = Int(Rnd * 255 + 1)
           g = Int(Rnd * 255 + 1)
           b = Int(Rnd * 255 + 1)
           Formzjm.Line (0, j)-(Me.Width, j), RGB(r, g, b)
       Next j
ElseIf nz = 3 Then
       For j = 1 To Me.Width
           r = Int(Rnd * 255 + 1)
           g = Int(Rnd * 255 + 1)
           b = Int(Rnd * 255 + 1)
           Formzjm.Line (j, 0)-(j, Me.Height), RGB(r, g, b)
       Next j
ElseIf nz = 1 Then
       Dim h%, z%, bj%, ng%
       For i = 1 To 100
           z = Rnd * Me.Height
           h = Rnd * Me.Width
           bj = Rnd * 3000
           r = Int(Rnd * 255 + 1)
           g = Int(Rnd * 255 + 1)
           b = Int(Rnd * 255 + 1)
           FillStyle = 0
           FillColor = RGB(r, g, b)
           ng = Int(Rnd * 3)
           If ng = 0 Then
              Formzjm.Circle (h, z), bj, RGB(r, g, b)
           ElseIf ng = 1 Then
                  Formzjm.Circle (h, z), bj, RGB(r, g, b), Rnd
           ElseIf ng = 2 Then
                  Formzjm.Line (h, z)-(h + bj, z + bj), RGB(r, g, b), BF
           End If
        Next i
ElseIf nz = 2 Then
       Dim sj!, jh%, jz%, sh%, sz%
       For i = 1 To 1000
           jz = Rnd * Me.Height
           jh = Rnd * Me.Width
           sz = Rnd * Me.Height
           sh = Rnd * Me.Width
           bj = Rnd * 2500
           r = Int(Rnd * 255 + 1)
           g = Int(Rnd * 255 + 1)
           b = Int(Rnd * 255 + 1)
           FillStyle = 0
           FillColor = RGB(r, g, b)
           sj = Rnd
           If sj < 0.55 Then
              Formzjm.Line (sh, jz)-(jh, sz), RGB(r, g, b)
           ElseIf sj < 0.65 Then
                  Formzjm.Line (jh, sz)-(sh + bj, jz + bj), RGB(r, g, b), BF
           ElseIf sj < 0.75 Then
                  Formzjm.Circle (jh, sz), bj, RGB(r, g, b)
           Else
               Formzjm.Circle (sh, jz), bj, RGB(r, g, b), Rnd
           End If
       Next i
End If
End Sub

Private Sub ssjs_Click()
ssjs.Checked = Not ssjs.Checked
If ssjs.Checked = False Then
   threej = ""
   xzjs(3) = ""
Else
    Call qxzt
    xzjs(3) = "����������"
End If
If md = 3 And laizou = False Then
   If Win(1).State = sckConnected And zhc = "bei" Then
      Win(1).SendData ("ssjs" & "js")
   End If
End If
End Sub

Private Sub tcbc_Click()
Call Combc_Click
End Sub

Private Sub tchq_Click()
Call Comhq_Click
End Sub

Private Sub tcks_Click()
If bsh >= 10 Or bsb >= 10 Then
   Call Comks_Click(1)
Else
    Call Comks_Click(0)
End If
End Sub

Private Sub tcyx_Click()
Dim aq%
aq = MsgBox("ȷ���˳���Ϸ��", 36, "��ʾ")
If aq = vbYes Then
   End
End If
End Sub

Private Sub textdh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Textdh <> "" Then
      If Win(1).State = sckConnected Then
         Win(1).SendData (Textdh & "dh")
         lisdh.Text = lisdh & vbCrLf & dl.mz & " " & Time
         lisdh.Text = lisdh & vbCrLf & "  ��" & Textdh
         Textdh.Text = ""
         Textdh.SetFocus
      End If
   End If
End If
End Sub

Private Sub Comby_Click()
Call bfys_Click
End Sub

Private Sub Comhq_Click()  '����
Dim he%, zo%
If md = 1 Then
   If (bsh = 1 And bsb = 0) Or (bsh = 0 And bsb = 1) Or (bsh = 0 And bsb = 0) Then
      MsgBox "��ʱ���ɻ���", 48, "��ʾ"
      Exit Sub
   End If
   If slz = True Then
      Timerh.Enabled = False
      Timerb.Enabled = True
      slz = False
      Call jstr(bz(bsb), he, zo)
      bz(bsb) = ""
      Imab(bsb).Visible = False
      wz(he, zo) = ""
      bsb = bsb - 1
      Labelbsb.Caption = "��" & bsb & "��"
      Call qxzt
   Else
      Timerh.Enabled = True
      Timerb.Enabled = False
      slz = True
      Call jstr(hz(bsh), he, zo)
      hz(bsh) = ""
      Imah(bsh).Visible = False
      wz(he, zo) = ""
      bsh = bsh - 1
      Labelbsh.Caption = "��" & bsh & "��"
      Call qxzt
   End If
   tsbc = False
   Picture2.Enabled = True
   If fzqz.Checked = False Then
   Call hqp
   For i = 0 To dhz - 1
       For j = 0 To dhz - 1
           If wz(i, j) = "����" Then
              For m = 1 To 100
                  Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
              Next m
           End If
           If wz(i, j) = "����" Then
              For l = 1 To 100
                  Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
              Next l
           End If
       Next j
   Next i
   Call fk(Not slz)
   End If
ElseIf md = 2 Then
       If (bsh <= 1 And bsb <= 0) Or (bsh <= 1 And bsb <= 1) Or (bsh = 0 And bsb = 0) Then
          MsgBox "��ʱ���ɻ���", 48, "��ʾ"
          Exit Sub
       End If
       Call jstr(bz(bsb), he, zo)
       bz(bsb) = ""
       Imab(bsb).Visible = False
       wz(he, zo) = ""
       bsb = bsb - 1
       Labelbsb.Caption = "��" & bsb & "��"
       '////////////////////////////
       If slz = False Then
          Call jstr(hz(bsh), he, zo)
          hz(bsh) = ""
          Imah(bsh).Visible = False
          wz(he, zo) = ""
          bsh = bsh - 1
          Labelbsh.Caption = "��" & bsh & "��"
          Call qxzt
       End If
       slz = False
       Timerb.Enabled = True
       Picture2.Enabled = True
       tsbc = False
       If fzqz.Checked = False Then
       Call hqp
       For i = 0 To dhz - 1
           For j = 0 To dhz - 1
               If wz(i, j) = "����" Then
                  For m = 1 To 100
                      Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
                  Next m
               End If
               If wz(i, j) = "����" Then
                  For l = 1 To 100
                      Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
                  Next l
               End If
           Next j
       Next i
       Call fk(Not slz)
       End If
ElseIf md = 3 Then
       If (bsh <= 1 And bsb <= 0) Or (bsh <= 0 And bsb <= 1) Or (bsh = 0 And bsb = 0) Then
          MsgBox "��ʱ���ɻ���", 48, "��ʾ"
          Exit Sub
       End If
       If slz = True Then
          MsgBox "��ʱ���ɻ���,�����ӣ�", 48, "��ʾ"
          Exit Sub
       End If
       If Win(1).State = sckConnected Then
          Win(1).SendData ("hq")
          Labip.Caption = "�ȴ��Է������������"
       End If
       Picture2.Enabled = False
End If
End Sub

Private Sub Comhy_Click()
Call hfys_Click
End Sub

Private Sub Comjl_Click()
If Comjl.Caption = "�� �� �� �� ��" Then
   Labip.Caption = "����IP��ַ:" & Win(0).LocalIP & "��" & "�ȴ��ͻ������ӡ�"
   Win(0).Bind 1992
   Win(0).Listen
   Comlj.Enabled = False
   Textip.Enabled = False
   jsxz.Enabled = False
   Comjl.Caption = "�� �� �� �� ��"
ElseIf Comjl.Caption = "�� �� �� �� ��" Then
       Win(0).Close
       Win(1).Close
       jsxz.Enabled = True
       Comlj.Enabled = True
       Textip.Enabled = True
       Comjl.Caption = "�� �� �� �� ��"
       Lal2.Caption = "����"
       Labip.Caption = "�������ѹرգ�"
       MsgBox "�����ѶϿ���������㣬���������Ӳ���ʼ��", 48, "��ʾ"
       Picture2.Enabled = True
       bsh = 0: bsb = 0
       Call Comks_Click(0)
       Picture2.Enabled = False
End If
End Sub

Private Sub Comks_Click(index As Integer)        '���¿�ʼ��Ϸ
If bsh >= 1 And bsb >= 1 Then
   If bt = True Then
      If tsbc = False Then
         Dim ad%
         ad = MsgBox("�Ƿ񱣴����ף�", 36, "��ʾ")
         If ad = vbYes Then
            Call Combc_Click
         End If
      End If
   End If
End If
If md = 3 And index = 1 Then
   If Win(1).State = sckConnected Then
      Win(1).SendData ("ks")
      Labip.Caption = "�ȴ��Է�ѡ���Ƿ�ʼ����Ϸ����"
   End If
   Exit Sub
End If
If ((index = 1 Or index = 2) And (bsh >= 10 Or bsb >= 10) And tsbc = False) Or (tsbc = False And index = 1 And bsh = 1 And bsb = 0) Then
   Dim yl As dlm
   Open App.Path & "zcb.lsn" For Random As #1 Len = Len(yl)
        For i = 1 To LOF(1) / Len(yl)
            Get #1, i, yl
            If yl.mz = dl.mz Then
               Exit For
            End If
        Next i
        If md = 1 Then
           yl.drh.bs_u = bsh + yl.drh.bs_u
           yl.drh.sj_u = sjh + yl.drh.sj_u
           yl.drb.bs_u = bsb + yl.drb.bs_u
           yl.drb.sj_u = sjb + yl.drb.sj_u
           yl.drh.undone = yl.drh.undone + 1
           yl.drb.undone = yl.drb.undone + 1
        ElseIf md = 2 Then
               yl.rj.bs_u = bsb + yl.rj.bs_u
               yl.rj.sj_u = sjb + yl.rj.sj_u
               yl.rj.undone = yl.rj.undone + 1
        ElseIf md = 3 Then
               yl.wl.bs_u = bsh + yl.wl.bs_u
               yl.wl.sj_u = sjh + yl.wl.sj_u
               yl.wl.undone = yl.wl.undone + 1
        End If
        Put #1, i, yl
        dl = yl
   Close #1
End If
Picture1.Enabled = True
Picture2.Enabled = True
Call hqp
If md = 3 Then
   Picture2.Enabled = False
End If
For i = 0 To 24
    For j = 0 To 24
        wz(i, j) = ""
    Next j
Next i
For i = 1 To 313
    Imah(i).Visible = False
    Imab(i).Visible = False
    hz(i) = ""
    bz(i) = ""
Next i
bsh = 0: bsb = 0: sjh = 0: sjb = 0
Timerh = False: Timerb = False
Labelsjh = "": Labelsjb = ""
Labelbsh = "": Labelbsb = ""
Frame1.Enabled = True
Frame1.Caption = "��Ϸ��ʼ����ѡ��"
If Lal2.Caption <> "����" Or md <> 3 Then
   If fzqz.Checked = False Then
      Comhy.Visible = True
   End If
End If
If md <> 3 Then
   If fzqz.Checked = False Then
      Comby.Visible = True
   End If
End If
If fzqz.Checked = False Then
   ys.Enabled = True
   tcys.Enabled = True
End If
dzms.Enabled = True
qpsz.Enabled = True
szqp.Enabled = True
zykj.Enabled = True
Option2.Value = True
lzd = "": qxh = "": qxb = ""
Labelts.Caption = ""
fourj = "": threej = ""
If index <> 0 Then
   Call qxxz_Click
   For i = 1 To 5
       xzjs(i) = ""
   Next i
End If
ys1 = RGB(0, 0, 0)
ys2 = RGB(255, 255, 255)
Pich.Scale (0, Pich.Height)-(Pich.Width, 0)
For i = 0 To Pich.Height / 2
    Pich.Circle (Pich.Width / 2, Pich.Width / 2), i, ys1
Next i
Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
For j = 0 To Picb.Height / 2
    Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, ys2
Next j
End Sub

Private Sub Comlj_Click()
If Comlj.Caption = "�� �� �� �� ��" Then
   Win(1).Connect Textip, 1992
   Comjl.Enabled = False
   Textip.Enabled = False
   Comlj.Caption = "�� �� �� ��"
   Labip.Caption = "�����С���"
ElseIf Comlj.Caption = "�� �� �� ��" Then
       Win(1).Close
       Comjl.Enabled = True
       Textip.Enabled = True
       Comlj.Caption = "�� �� �� �� ��"
       Lal2.Caption = "����"
       Labip.Caption = "�����ѶϿ������������ӣ�"
       MsgBox "�����ѶϿ���������㣬���������Ӳ���ʼ��", 48, "��ʾ"
       Picture2.Enabled = True
       bsh = 0: bsb = 0
       Call Comks_Click(0)
       Picture2.Enabled = False
End If
End Sub

Private Sub drdz_Click()
md = 1
drdz.Checked = True
drdz.Enabled = False
rjdz.Enabled = True
rjdz.Checked = False
wldz.Enabled = True
wldz.Checked = False
Call kjbj
If Win(0).State <> sckClosed Then
   Win(0).Close
End If
If Win(1).State <> sckClosed Then
   Win(1).Close
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If (bsh >= 1 And bsb >= 1) Then
   If tsbc = False Then
      Dim aq%
      aq = MsgBox("�Ƿ��ڹر�ǰ�������ף�", 36, "��ʾ")
      If aq = vbYes Then
         Call Combc_Click
         If UnloadMode = 0 Then
            Cancel = 1
         ElseIf UnloadMode = 1 Then
                Cancel = 0
         End If
      End If
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Win(0).State <> sckClosed Then
   Win(0).Close
End If
If Win(1).State <> sckClosed Then
   Win(1).Close
End If
End Sub

Private Sub rjdz_Click()
md = 2
rjdz.Checked = True
rjdz.Enabled = False
drdz.Enabled = True
drdz.Checked = False
wldz.Enabled = True
wldz.Checked = False
Call kjbj
If Win(0).State <> sckClosed Then
   Win(0).Close
End If
If Win(1).State <> sckClosed Then
   Win(1).Close
End If
End Sub

Private Sub scqp_Click()  '�鿴�ϴ��������Ϸ������
If bsh >= 1 Or bsb >= 1 Then
   If bt = True Then
      If tsbc = False Then
         Dim ad%
         ad = MsgBox("�Ƿ񱣴����ף�", 36, "��ʾ")
         If ad = vbYes Then
            Call Combc_Click
         End If
      End If
   End If
End If
If bsh >= 10 Or bsb >= 10 Then
   bsh = 1: bsb = 0
   Call Comks_Click(1)
Else
    bsh = 0: bsb = 0
    Call Comks_Click(0)
End If
tsbc = True
Comhy.Visible = False
Comby.Visible = False
ys.Enabled = False
tcys.Enabled = False
dzms.Enabled = False
qpsz.Enabled = False
szqp.Enabled = False
zykj.Enabled = False
Dim cv As save, jlh%, h%, z%, send$, bshb$, jr$, yj$
   Open App.Path & "���������.lsl" For Random As #2 Len = Len(cv)
        FileName = "": jlh = 1
        Do Until LOF(2) / Len(cv) < jlh
           Get #2, jlh, cv
           If jlh = 1 Then
              sjh = cv.sjh: sjb = cv.sjb
              ys1 = cv.ysh: ys2 = cv.ysb
              Pich.Scale (0, Pich.Height)-(Pich.Width, 0)
              For i = 0 To Pich.Height / 2
                  Pich.Circle (Pich.Width / 2, Pich.Width / 2), i, ys1
              Next i
              Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
              For j = 0 To Picb.Height / 2
                  Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, ys2
              Next j
              If md = 1 Then
                 Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
                 If sjh = 0 Then
                    sjh = sjb
                    Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                 Else
                     Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                 End If
              ElseIf md = 2 Then
                     Labelsjb.Caption = dl.mz & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
              ElseIf md = 3 Then
                     Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
                     If sjh = 0 Then
                        sjh = sjb
                        Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                     Else
                         Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
                     End If
                     If Win(1).State = sckConnected Then
                        send = "," & sjh & "," & sjb & "," & ys1 & "," & ys2
                     End If
              End If
           End If
           Select Case Trim(cv.zbh)
                   Case ""
                        If cv.zbb = "1234" Then
                           bsh = jlh - 1: bsb = jlh - 1
                           Labelbsb.Caption = "��" & bsb & "��"
                           Labelbsh.Caption = "��" & bsh & "��"
                           If Win(1).State = sckConnected Then
                              bshb = "," & bsh & "," & bsb
                           End If
                           If md <> 2 Then
                              Timerh.Enabled = True
                           End If
                           slz = True
                           Option1.Value = True
                           Frame1.Caption = "��Ϸ�У�����ѡ��"
                           Frame1.Enabled = False
                           Call jstr(Trim(Str(cv.sjb)), h, z)
                           If z > dhz Then
                              If Win(1).State = sckConnected Then
                                 jr = z
                              End If
                              MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                              bsh = 0: bsb = 0
                              Call Comks_Click(0)
                              slz = True
                              Picture2.Enabled = True
                              Exit Do
                           Else
                               jr = "0"
                           End If
                           If h = 1 Then
                              Timerh.Enabled = False
                              Picture2.Enabled = False
                              If Win(1).State = sckConnected Then
                                 yj = "b"
                                 MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                              Else
                                  MsgBox "������" & " " & "�׷�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                              End If
                           ElseIf h = 2 Then
                                  Timerh.Enabled = False
                                  Picture2.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & "���" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                           ElseIf h = 3 Then
                                  Timerh.Enabled = False
                                  Picture2.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                           ElseIf md = 3 Then
                                  Timerh.Enabled = True
                                  If Win(1).State = sckConnected Then
                                     yj = "bs"
                                  End If
                           End If
                        ElseIf Trim(cv.zbb) <> "" Then
                               bsh = jlh - 1: bsb = jlh
                               Labelbsb.Caption = "��" & bsb & "��"
                               Labelbsh.Caption = "��" & bsh & "��"
                               If Win(1).State = sckConnected Then
                                  bshb = "," & bsh & "," & bsb
                               End If
                               If md <> 2 Then
                                  Timerh.Enabled = True
                               End If
                               slz = True
                               Option2.Value = True
                               Frame1.Caption = "��Ϸ�У�����ѡ��"
                               Frame1.Enabled = False
                               bz(jlh) = cv.zbb
                               Call jstr(cv.zbb, h, z)
                               wz(h, z) = "����"
                               Call hqz(h, z, bsb, Not slz)
                               If Win(1).State = sckConnected Then
                                  send = send + "," & cv.zbb & jlh & "h"
                               End If
                               Call jstr(Trim(Str(cv.sjb)), h, z)
                               If z > dhz Then
                                  If Win(1).State = sckConnected Then
                                     jr = z
                                  End If
                                  MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                                  bsh = 0: bsb = 0
                                  Call Comks_Click(0)
                                  slz = True
                                  Picture2.Enabled = True
                                  Exit Do
                               Else
                                   jr = "0"
                               End If
                               If h = 1 Then
                                  Picture2.Enabled = False
                                  Timerh.Enabled = False
                                  If Win(1).State = sckConnected Then
                                     yj = "b"
                                     MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                  Else
                                      MsgBox "������" & " " & "�׷�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                  End If
                               ElseIf h = 2 Then
                                      Picture2.Enabled = False
                                      Timerh.Enabled = False
                                      If Win(1).State = sckConnected Then
                                         yj = "b"
                                         MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                      Else
                                          MsgBox "������" & " " & "���" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                      End If
                               ElseIf h = 3 Then
                                      Picture2.Enabled = False
                                      Timerh.Enabled = False
                                      If Win(1).State = sckConnected Then
                                         yj = "b"
                                         MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壬�����ӣ�", 48, "��ʾ"
                                      Else
                                          MsgBox "������" & " " & Lal2 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                      End If
                               ElseIf md = 3 Then
                                      Timerh.Enabled = True
                                      If Win(1).State = sckConnected Then
                                         yj = "bs"
                                      End If
                               End If
                        End If
                   Case "1234"
                        bsb = jlh - 1: bsh = jlh - 1
                        Labelbsb.Caption = "��" & bsb & "��"
                        Labelbsh.Caption = "��" & bsh & "��"
                        If Win(1).State = sckConnected Then
                           bshb = "," & bsh & "," & bsb
                        End If
                        Timerb.Enabled = True
                        slz = False
                        Option2.Value = True
                        Frame1.Caption = "��Ϸ�У�����ѡ��"
                        Frame1.Enabled = False
                        Call jstr(Trim(Str(cv.sjh)), h, z)
                        If z > dhz Then
                           If Win(1).State = sckConnected Then
                              jr = z
                           End If
                           MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                           bsh = 0: bsb = 0
                           Call Comks_Click(0)
                           slz = True
                           Picture2.Enabled = True
                           Exit Do
                        Else
                            jr = "0"
                        End If
                        If h = 1 Then
                           Timerb.Enabled = False
                           Picture2.Enabled = False
                           If Win(1).State = sckConnected Then
                              yj = "h"
                              MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                           Else
                               MsgBox "������" & " " & "�ڷ�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                           End If
                        ElseIf h = 2 Then
                               Timerb.Enabled = False
                               Picture2.Enabled = False
                               If Win(1).State = sckConnected Then
                                  yj = "h"
                                  MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                               Else
                                  MsgBox "������" & " " & "����" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                               End If
                        ElseIf h = 3 Then
                               Timerb.Enabled = False
                               Picture2.Enabled = False
                               If Win(1).State = sckConnected Then
                                  yj = "h"
                                  MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                               Else
                                   MsgBox "������" & " " & Lal1 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                               End If
                        ElseIf md = 3 Then
                               Timerb.Enabled = True
                               If Win(1).State = sckConnected Then
                                  yj = "hs"
                               End If
                        End If
                   Case Else
                       If Trim(cv.zbb) = "" Then
                          bsh = jlh: bsb = jlh - 1
                          Labelbsb.Caption = "��" & bsb & "��"
                          Labelbsh.Caption = "��" & bsh & "��"
                          If Win(1).State = sckConnected Then
                             bshb = "," & bsh & "," & bsb
                          End If
                          Timerb.Enabled = True
                          slz = False
                          Option1.Value = True
                          Frame1.Caption = "��Ϸ�У�����ѡ��"
                          Frame1.Enabled = False
                          hz(jlh) = cv.zbh
                          Call jstr(cv.zbh, h, z)
                          wz(h, z) = "����"
                          Call hqz(h, z, bsh, Not slz)
                          If Win(1).State = sckConnected Then
                             send = send + "," & cv.zbh & jlh & "b"
                          End If
                          Call jstr(Trim(Str(cv.sjh)), h, z)
                          If z > dhz Then
                              If Win(1).State = sckConnected Then
                                 jr = z
                              End If
                              MsgBox "�����̲����ݴ�����(" & z & "��" & z & ")�����黻һ����������!", 48, "��ʾ"
                              bsh = 0: bsb = 0
                              Call Comks_Click(0)
                              slz = True
                              Picture2.Enabled = True
                              Exit Do
                          Else
                              jr = "0"
                          End If
                          If h = 1 Then
                             Timerb.Enabled = False
                             Picture2.Enabled = False
                             If Win(1).State = sckConnected Then
                                yj = "h"
                                MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                             Else
                                 MsgBox "������" & " " & "�ڷ�" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                             End If
                          ElseIf h = 2 Then
                                 Timerb.Enabled = False
                                 Picture2.Enabled = False
                                 If Win(1).State = sckConnected Then
                                    yj = "h"
                                    MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 Else
                                     MsgBox "������" & " " & "����" & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 End If
                          ElseIf h = 3 Then
                                 Timerb.Enabled = False
                                 Picture2.Enabled = False
                                 If Win(1).State = sckConnected Then
                                    yj = "h"
                                    MsgBox "���������ѻ�ʤ" & "�������Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 Else
                                    MsgBox "������" & " " & Lal1 & " " & "�ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
                                 End If
                          ElseIf md = 3 Then
                                 Timerb.Enabled = True
                                 If Win(1).State = sckConnected Then
                                    yj = "hs"
                                 End If
                          End If
                       Else
                           hz(jlh) = cv.zbh
                           Call jstr(cv.zbh, h, z)
                           wz(h, z) = "����"
                           Call hqz(h, z, jlh, tsbc)
                           If Win(1).State = sckConnected Then
                              send = send + "," & cv.zbh & jlh & "b"
                           End If
                           '///////////////////////////////////
                           bz(jlh) = cv.zbb
                           Call jstr(cv.zbb, h, z)
                           wz(h, z) = "����"
                           Call hqz(h, z, jlh, Not tsbc)
                           If Win(1).State = sckConnected Then
                              send = send + "," & cv.zbb & jlh & "h"
                           End If
                       End If
           End Select
           jlh = jlh + 1
        Loop
   Close #2
   If Win(1).State = sckConnected Then
      send = jr & bshb & "," & yj & send & "lp"
      Win(1).SendData send
   End If
End Sub

Private Sub sdys_Click()
com.CancelError = True
On Error GoTo errhandler
com.ShowColor
   HS.Visible = False
   VS.Visible = False
   Picture1.Picture = LoadPicture("")
   Picture1.BackColor = com.Color
   Picture1.Height = Picture2.Height
   Picture1.Width = Picture2.Width
   Picture1.Top = -15
   Picture1.Left = -15
   Picture1.Scale (0, dhz * 10)-(dhz * 10, 0)
   Call hqp
   If fzqz.Checked = False Then
For i = 0 To 24
    For j = 0 To 24
        If wz(i, j) = "����" Then
           For m = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
           Next m
        End If
        If wz(i, j) = "����" Then
           For l = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
           Next l
        End If
    Next j
Next i
If bsb >= 1 Or bsh >= 1 Then
   If slz = True Then
      Call fk(Not slz)
   ElseIf slz = False Then
          Call fk(Not slz)
   End If
End If
End If
errhandler:
End Sub

Private Sub tcdl_Click()
Dim ad%
ad = MsgBox("�Ƿ��˳������µ�¼��", 36, "��ʾ")
If ad = vbYes Then
   Call Comks_Click(0)
   Unload Me
   Formsd.Show
End If
End Sub

Private Sub Timersj_Timer()  '��ʾ��ǰʱ��
Labelsj.Caption = Now
Static bbca$, cs%, sz As Boolean
Dim sjs!
cs = cs + 1
If cs = 20 And sz = False Then
   cs = 0
   sz = True
   bbca = Labip.Caption
   Randomize
   sjs = Int(Rnd * 43)
   Labip = smq(sjs)
ElseIf cs = 40 And sz = True Then
       cs = 0
       Labip = bbca
       sz = False
End If
If sz = False Then
   bbca = Labip
End If
End Sub

Private Sub tjtp_Click()
com.CancelError = True
On Error GoTo errhandler
com.ShowOpen
If Right(com.FileName, 3) <> "JPG" And Right(com.FileName, 3) <> "BMP" And Right(com.FileName, 3) <> "GIF" And Right(com.FileName, 3) <> "jpg" And Right(com.FileName, 3) <> "bmp" And Right(com.FileName, 3) <> "gif" Then
   MsgBox "����غ�׺��Ϊ��jpg,bmp��gif����ͼƬ", 48, "��ʾ"
   Exit Sub
End If
Picture1.Top = -15
Picture1.Left = -15
Picture1.Height = Picture2.Height
Picture1.Width = Picture2.Width
Picture1.Scale (0, dhz * 10)-(dhz * 10, 0)
Picture1 = LoadPicture(com.FileName)
If Picture1.Width < Picture2.Width And Picture1.Height < Picture2.Height Then
   HS.Visible = False
   VS.Visible = False
   Picture1.Height = Picture2.Height
   Picture1.Width = Picture2.Width
End If
If Picture1.Height > Picture2.Height Then
   VS.Visible = True
   VS.Height = Picture2.Height
   If Picture1.Height - Picture2.Height > 32767 Then
      VS.Max = 32767
      v1 = (Picture1.Height - Picture2.Height) / 32767
   Else
       v1 = 0
       VS.Max = Picture1.Height - Picture2.Height
   End If
   VS.SmallChange = VS.Max / 20
   VS.LargeChange = VS.Max / 20
   VS.Value = 0
Else
   Picture1.Height = Picture2.Height
   VS.Visible = False
End If
If Picture1.Width > Picture2.Width Then
   HS.Visible = True
   HS.Width = Picture2.Width
   If Picture1.Width - Picture2.Width > 32767 Then
      HS.Max = 32767
      h1 = (Picture1.Width - Picture2.Width) / 32767
   Else
       h1 = 0
       HS.Max = Picture1.Width - Picture2.Width
   End If
   HS.SmallChange = HS.Max / 20
   HS.LargeChange = HS.Max / 20
   HS.Value = 0
Else
   Picture1.Width = Picture2.Width
   HS.Visible = False
End If
Call hqp
If fzqz.Checked = False Then
For i = 0 To 24
    For j = 0 To 24
        If wz(i, j) = "����" Then
           For m = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
           Next m
        End If
        If wz(i, j) = "����" Then
           For l = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
           Next l
        End If
    Next j
Next i
If bsb >= 1 Or bsh >= 1 Then
   If slz = True Then
      Call fk(Not slz)
   ElseIf slz = False Then
          Call fk(Not slz)
   End If
End If
End If
errhandler:
End Sub

Private Sub Tmrxz_Timer()
For i = 1 To 5
    If xzjs(i) <> "" Then
       xzz = xzz & "-" & xzjs(i)
    End If
Next i
If xzz <> "" Then
   xzz = Right(xzz, Len(xzz) - 1)
End If
Labelxz.Caption = xzz
If xzz <> "" Then
   If Picxz.Width < Labelxz.Width Then
      Labelxz.Left = Labelxz.Left - 50
      If Labelxz.Left <= Picxz.Width - Labelxz.Width - 300 Then
         Labelxz.Left = 300
      End If
   End If
End If
If Picip.Width < Labip.Width Then
   Labip.Left = Labip.Left - 50
   If Labip.Left <= Picip.Width - Labip.Width - 400 Then
      Labip.Left = 400
   End If
Else
    Labip.Left = 0
End If
End Sub

Private Sub tupian_Click()
Call tjtp_Click
End Sub

Private Sub VS_Change()
If v1 = 0 Then
   Picture1.Top = -VS.Value
   Call chqp
Else
    Picture1.Top = -VS.Value * v1
    Call chqp
End If
End Sub

Private Sub VS_Scroll()
If v1 = 0 Then
   Picture1.Top = -VS.Value
   Call chqp
Else
    Picture1.Top = -VS.Value * v1
    Call chqp
End If
End Sub
Private Sub HS_Change()
If h1 = 0 Then
   Picture1.Left = -HS.Value
   Call chqp
Else
    Picture1.Left = -HS.Value * h1
    Call chqp
End If
End Sub
Private Sub HS_Scroll()
If h1 = 0 Then
   Picture1.Left = -HS.Value
   Call chqp
Else
    Picture1.Left = -HS.Value * h1
    Call chqp
End If
End Sub

Private Sub Win_Close(index As Integer)
If index = 1 Then
   Labip.Caption = "�����ѶϿ���"
   Comhy.Visible = False
   Comjl.Enabled = True
   Comjl.Caption = "�� �� �� �� ��"
   Comlj.Enabled = True
   Comlj.Caption = "�� �� �� �� ��"
   Textip.Enabled = True
   Lal2.Caption = "����"
   Win(0).Close
   Win(1).Close
   Picture2.Enabled = True
   bsh = 0: bsb = 0
   Call Comks_Click(0)
   Picture2.Enabled = False
End If
End Sub

Private Sub Win_Connect(index As Integer)
If Win(index).State = sckConnected Then
   Win(index).SendData (dl.mz & "kh")
   Labip.Caption = "�����ӡ���"
ElseIf Win(index).State = sckClosed Then
       Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
End If
End Sub

Private Sub Win_ConnectionRequest(index As Integer, ByVal requestID As Long)
Win(1).Accept requestID
Labip.Caption = "�����ӡ���"
If Win(1).State = sckConnected Then
   Win(1).SendData (dl.mz & "fw")
ElseIf Win(1).State = sckClosed Then
       Win(1).Connect Win(1).RemoteHostIP, Win(1).RemotePort
End If
End Sub

Private Sub Win_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim sda$, li$, se$, mg%, he%, zo%
Win(index).GetData sda, vbString
li = Right(sda, 2)
se = Mid(sda, 1, Len(sda) - 2)
If li = "fw" Then
   zhc = "bei"
   Lal2.Caption = se
   bfys.Caption = se & "��ɫ"
   bye.Caption = se & "��ɫ"
   qpsz.Enabled = False
   szqp.Enabled = False
   jsxz.Enabled = True
   zykj.Enabled = False
   yxxz.Enabled = True
   ckbc.Enabled = True
   scqp.Enabled = False
   If fzqz.Checked = False Then
      Comhy.Visible = True
      ys.Enabled = True
      tcys.Enabled = True
      hfys.Enabled = True
      bfys.Enabled = False
      bye.Enabled = False
      hy.Enabled = True
   End If
ElseIf li = "kh" Then
       Lal2.Caption = se
       bfys.Caption = se & "��ɫ"
       bye.Caption = se & "��ɫ"
       yxxz.Enabled = True
       ckbc.Enabled = True
       If bsh = 0 And bsb = 0 Then
          mg = MsgBox("�ɹ��ͶԷ����ӣ��Ƿ񼺷��������ӣ�", 36, "�����ķ�����")
          If mg = vbYes Then
             Picture2.Enabled = True
             slz = True
             zhc = "zhu"
             szqp.Enabled = True
             qpsz.Enabled = True
             jsxz.Enabled = False
             zykj.Enabled = True
             scqp.Enabled = True
             If fzqz.Checked = False Then
                Comhy.Visible = True
                ys.Enabled = True
                tcys.Enabled = True
                hfys.Enabled = True
                bfys.Enabled = False
                bye.Enabled = False
                hy.Enabled = True
             End If
          ElseIf mg = vbNo Then
                 zhc = "bei"
                 qpsz.Enabled = False
                 szqp.Enabled = False
                 jsxz.Enabled = True
                 zykj.Enabled = False
                 scqp.Enabled = False
                 If fzqz.Checked = False Then
                    Comhy.Visible = True
                    ys.Enabled = True
                    tcys.Enabled = True
                    hfys.Enabled = True
                    bfys.Enabled = False
                    bye.Enabled = False
                    hy.Enabled = True
                End If
                If Win(index).State = sckConnected Then
                   Win(index).SendData ("��������" & "lz")
                ElseIf Win(index).State = sckClosed Then
                       Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
                End If
          End If
       End If
ElseIf li = "lz" Then
       MsgBox se, 0, "�Է���Ϣ"
       zhc = "zhu"
       Picture2.Enabled = True
       If fzqz.Checked = False Then
          Comhy.Visible = True
       End If
       slz = True
       szqp.Enabled = True
       qpsz.Enabled = True
       jsxz.Enabled = False
       zykj.Enabled = True
       scqp.Enabled = True
       If fzqz.Checked = False Then
          Comhy.Visible = True
          ys.Enabled = True
          tcys.Enabled = True
          hfys.Enabled = True
          bfys.Enabled = False
          bye.Enabled = False
          hy.Enabled = True
       End If
ElseIf li = "ys" Then
       ys2 = Val(se)
       Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
       For j = 0 To Picb.Height / 2
           Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, ys2
       Next j
ElseIf li = "wz" Then
       Call jstr(se, hzb, zzb)
       If slz = False Then
          tcys.Enabled = False
          ys.Enabled = False
          Comhy.Visible = False
          dzms.Enabled = False
          qpsz.Enabled = False
          szqp.Enabled = False
          zykj.Enabled = False
          If ys1 = 0 And ys2 = 0 Then
             ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
          End If
          Picture2.Enabled = True
          Labelzb.Caption = Chr(65 + hzb) & zzb + 1 & " " & "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
          lzd = "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
          Call hqz(hzb, zzb, bsb + 1, slz)
          slz = True
          wz(hzb, zzb) = "����"
          bsb = bsb + 1
          bz(bsb) = bstr(hzb, zzb)
          Labelbsb.Caption = "��" & bsb & "��"
          If fzqz.Checked = False Then
             Call fk(Not slz)
          End If
          tsbc = False
          Timerb.Enabled = False
          Timerh.Enabled = True
          Call qxzt
          If pdwz(wz(), hzb, zzb, Lal2.Caption) = True Then
             Exit Sub
          End If
       End If
ElseIf li = "hq" Then
       Timerh.Enabled = False
       mg = MsgBox("�Է�������壬�Ƿ���ܣ�", 36, "�Է���Ϣ")
       If mg = vbYes Then
          If Win(index).State = sckConnected Then
             Win(index).SendData ("ty")
          ElseIf Win(index).State = sckClosed Then
                 Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
          End If
          Timerh.Enabled = False
          Timerb.Enabled = True
          slz = False
          Call jstr(bz(bsb), he, zo)
          bz(bsb) = ""
          Imab(bsb).Visible = False
          wz(he, zo) = ""
          bsb = bsb - 1
          Labelbsb.Caption = "��" & bsb & "��"
          Call qxzt
          If fzqz.Checked = False Then
             Call hqp
             For i = 0 To 24
                 For j = 0 To 24
                     If wz(i, j) = "����" Then
                        For m = 1 To 100
                            Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
                        Next m
                     End If
                     If wz(i, j) = "����" Then
                        For l = 1 To 100
                            Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
                        Next l
                     End If
                 Next j
             Next i
             Call fk(Not slz)
          End If
       ElseIf mg = vbNo Then
              Timerh.Enabled = True
              If Win(index).State = sckConnected Then
                 Win(index).SendData ("no")
              ElseIf Win(index).State = sckClosed Then
                     Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
              End If
       End If
ElseIf li = "ty" Then
       MsgBox "ͬ����Ļ��壡", 0, "�Է���Ϣ"
       Labip.Caption = "�����ӡ���"
       Picture2.Enabled = True
       slz = True
       Timerh.Enabled = True
       Timerb.Enabled = False
       Call jstr(hz(bsh), he, zo)
       hz(bsh) = ""
       Imah(bsh).Visible = False
       wz(he, zo) = ""
       bsh = bsh - 1
       Labelbsh.Caption = "��" & bsh & "��"
       Call qxzt
       If fzqz.Checked = False Then
          Call hqp
          For i = 0 To 24
              For j = 0 To 24
                  If wz(i, j) = "����" Then
                     For m = 1 To 100
                         Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
                     Next m
                  End If
                  If wz(i, j) = "����" Then
                     For l = 1 To 100
                         Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
                     Next l
                  End If
              Next j
          Next i
          Call fk(Not slz)
       End If
ElseIf li = "no" Then
       Timerh.Enabled = False
       MsgBox "��ͬ����Ļ��壡", 48, "�Է���Ϣ"
       Labip.Caption = "�����ӡ���"
       Picture2.Enabled = True
       Timerb.Enabled = True
       Timerh.Enabled = False
ElseIf li = "ks" Then
       Timerh.Enabled = False
       mg = MsgBox("�Է��������¿�ʼ��Ϸ���Ƿ���ܣ�", 36, "�Է���Ϣ")
       If mg = vbYes Then
          Picture2.Enabled = True
          If bsh >= 10 Or bsb >= 10 Then
             Call Comks_Click(2)
          Else
              Call Comks_Click(0)
          End If
          Picture2.Enabled = False
          If Win(index).State = sckConnected Then
             Dim ra!
             Randomize
             ra = Rnd
             If Round(ra) = 1 Then
                Win(index).SendData ("nizou" & "yt")
                MsgBox "���ѡ�������Է������ӣ�", 48, "��ʾ"
                slz = False
                zhc = "bei"
                szqp.Enabled = False
                qpsz.Enabled = False
                jsxz.Enabled = True
                Comhy.Visible = False
                zykj.Enabled = False
                scqp.Enabled = False
             Else
                 Win(index).SendData ("wozou" & "yt")
                 MsgBox "���ѡ���������������ӣ�", 48, "��ʾ"
                 Picture2.Enabled = True
                 slz = True
                 zhc = "zhu"
                 szqp.Enabled = True
                 qpsz.Enabled = True
                 jsxz.Enabled = False
                 zykj.Enabled = True
                 scqp.Enabled = True
             End If
             If fzqz.Checked = False Then
                Comhy.Visible = True
                ys.Enabled = True
                tcys.Enabled = True
                hfys.Enabled = True
                bfys.Enabled = False
                bye.Enabled = False
                hy.Enabled = True
             End If
          ElseIf Win(index).State = sckClosed Then
                 Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
          End If
       ElseIf mg = vbNo Then
              Timerh.Enabled = True
              If Win(index).State = sckConnected Then
                 Win(index).SendData ("nb")
              ElseIf Win(index).State = sckClosed Then
                     Win(index).Connect Win(index).RemoteHostIP, Win(index).RemotePort
              End If
       End If
ElseIf li = "yt" Then
       MsgBox "�Է���ͬ�����¿�ʼ��Ϸ", 48, "�Է���Ϣ"
       Labip.Caption = "�����ӡ���"
       Picture2.Enabled = True
       If bsh >= 10 Or bsb >= 10 Then
          Call Comks_Click(2)
       Else
           Call Comks_Click(0)
       End If
       Picture2.Enabled = False
       If se = "nizou" Then
          MsgBox "���ѡ���������������ӣ�", 48, "��ʾ"
          Picture2.Enabled = True
          slz = True
          zhc = "zhu"
          szqp.Enabled = True
          qpsz.Enabled = True
          jsxz.Enabled = False
          zykj.Enabled = True
          scqp.Enabled = True
       ElseIf se = "wozou" Then
              MsgBox "���ѡ�������Է������ӣ�", 48, "��ʾ"
              slz = False
              zhc = "bei"
              szqp.Enabled = False
              qpsz.Enabled = False
              jsxz.Enabled = True
              Comhy.Visible = False
              zykj.Enabled = False
              scqp.Enabled = False
       End If
       If fzqz.Checked = False Then
          Comhy.Visible = True
          ys.Enabled = True
          tcys.Enabled = True
          hfys.Enabled = True
          bfys.Enabled = False
          bye.Enabled = False
          hy.Enabled = True
       End If
ElseIf li = "nb" Then
       MsgBox "�Է���ͬ�����¿�ʼ��Ϸ��", 48, "�Է���Ϣ"
       Labip.Caption = "�����ӡ���"
ElseIf li = "dh" Then
       lisdh.Text = lisdh & vbCrLf & Lal2.Caption & " " & Time
       lisdh.Text = lisdh & vbCrLf & "  ��" & se
ElseIf li = "zz" Then
       zhc = se & zhc
       If Len(se) = 2 Then
          mg = Val((Right(se, 1)))
       ElseIf Len(se) = 3 Then
              mg = Val((Right(se, 2)))
       End If
       Call zz_Click(mg)
ElseIf li = "xz" Then
       zhc = se & zhc
       If Len(se) = 2 Then
          mg = Val((Right(se, 1)))
       ElseIf Len(se) = 3 Then
              mg = Val((Right(se, 2)))
       End If
       Call xz_Click(mg)
ElseIf li = "qp" Then
       Picture2.Enabled = True
       Call pq_Click(Val(se))
       Picture2.Enabled = False
ElseIf li = "xs" Then
       mg = Val(Right(se, 1))
       laizou = True
       If mg = 7 Then
          Call shi_Click(mg)
          mg = Val(Left(se, Len(se) - 1))
          xzs = mg
          xzjs(1) = "��ʱ" & xzs & "��"
       ElseIf mg = 8 Then
              Call shi_Click(mg)
       ElseIf mg >= 1 And mg <= 6 Then
              Call shi_Click(mg)
       End If
       laizou = False
ElseIf li = "xb" Then
       mg = Val(Right(se, 2))
       laizou = True
       If mg = 15 Then
          Call bu_Click(mg)
          mg = Val(Left(se, Len(se) - 2))
          xzb = mg
          xzjs(2) = "��" & xzb & "��"
       ElseIf mg = 16 Then
              Call bu_Click(mg)
       ElseIf mg > 16 Then
              mg = mg - 90
              Call bu_Click(mg)
       Else
           Call bu_Click(mg)
       End If
       laizou = False
ElseIf li = "js" Then
       laizou = True
       If se = "ssjs" Then
          Call ssjs_Click
       ElseIf se = "sijs" Then
              Call sijs_Click
       ElseIf se = "cljs" Then
              Call cljs_Click
       ElseIf li = "qxjs" Then
              Call qxjs_Click
       End If
       laizou = False
ElseIf li = "qx" Then
       laizou = True
       Call bu_Click(16)
       Call shi_Click(8)
       If zhc = "zhu" Then
          Call qxjs_Click
       End If
       laizou = False
ElseIf li = "lp" Then
       hu = Split(se, ",")
       If hu(0) > "0" Then
          MsgBox Lal2 & "�������ף��������̲����ݴ�����(" & hu(0) & "��" & hu(0) & ")�����黻һ����������!", 48, "��ʾ"
          bsb = 0: bsh = 0
          Call Comks_Click(0)
          Exit Sub
       End If
       bsh = Val(hu(2))
       bsb = Val(hu(1))
       Labelbsb.Caption = "��" & bsb & "��"
       Labelbsh.Caption = "��" & bsh & "��"
       sjb = Val(hu(4))
       Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
       sjh = Val(hu(5))
       Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
       ys2 = Val(hu(7))
       Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
       For j = 0 To Picb.Height / 2
           Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, ys2
       Next j
       ys1 = Val(hu(6))
       Pich.Scale (0, Pich.Height)-(Pich.Width, 0)
       For j = 0 To Pich.Height / 2
           Pich.Circle (Pich.Width / 2, Pich.Width / 2), j, ys1
       Next j
       Picture2.Enabled = True
       For i = 8 To UBound(hu)
           If Right(hu(i), 1) = "h" Then
              li = Left(hu(i), 4)
              mg = Val(Right(hu(i), Len(hu(i)) - 4))
              hz(mg) = li
              Call jstr(li, he, zo)
              wz(he, zo) = "����"
              Call hqz(he, zo, mg, True)
           ElseIf Right(hu(i), 1) = "b" Then
                  li = Left(hu(i), 4)
                  mg = Val(Right(hu(i), Len(hu(i)) - 4))
                  bz(mg) = li
                  Call jstr(li, he, zo)
                  wz(he, zo) = "����"
                  Call hqz(he, zo, mg, False)
           End If
       Next i
       Picture2.Enabled = False
       If hu(3) = "b" Then
          Picture2.Enabled = False
          slz = False
          MsgBox Lal2 & "�������ף����������ѻ�ʤ" & "�����Ȼ��壬�����ӣ�", 48, "��ʾ"
       ElseIf hu(3) = "h" Then
              Picture2.Enabled = False
              slz = True
              MsgBox Lal2 & "�������ף�������" & " " & Lal2 & " " & "�ѻ�ʤ" & "����ȴ��Է����壡", 48, "��ʾ"
       ElseIf hu(3) = "bs" Then
              Picture2.Enabled = True
              slz = True
              Timerh.Enabled = True
              MsgBox Lal2 & "�������ף������״˲��������ӣ������ӣ�", 48, "��ʾ"
       ElseIf hu(3) = "hs" Then
              slz = False
              Timerb.Enabled = True
              Picture2.Enabled = False
              MsgBox Lal2 & "�������ף������״˲��ɶԷ����ӣ���ȴ���", 48, "��ʾ"
       End If
End If
End Sub


Private Sub Win_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Win(index).State <> sckClosed Then
   Win(index).Close
End If
End Sub

Private Sub wldz_Click()
md = 3
rjdz.Enabled = True
rjdz.Checked = False
drdz.Enabled = True
drdz.Checked = False
wldz.Checked = True
wldz.Enabled = False
Call kjbj
If Win(0).State <> sckClosed Then
   Win(0).Close
End If
If Win(1).State <> sckClosed Then
   Win(1).Close
End If
End Sub

Private Sub wzqjj_Click()
ave = 1
Formsm.Show 1
End Sub

Private Sub xdtp_Click()
Call Image1_Click(inde)
End Sub

Private Sub xgmm_Click()
Dim mm$, dr As Boolean
dr = False
Do
  mm = Trim(InputBox("�����������룺" & vbCrLf & vbCrLf & "(�����10���ַ�,ֻ��Ϊ���ֻ�Ӣ��)", "�޸�����"))
  If Len(mm) > 10 Then
     MsgBox "�����������" & vbCrLf & "�����10���ַ�,ֻ��Ϊ���ֻ�Ӣ��", 48, "��ʾ"
  ElseIf Len(mm) = 0 Then
         MsgBox "�����޸�ʧ�ܣ�", 0, "��ʾ"
         Exit Sub
  Else
      For i = 1 To Len(mm)
          If (Asc(Mid(mm, i, 1)) > 47 And Asc(Mid(mm, i, 1)) < 58) Or _
          (Asc(Mid(mm, i, 1)) > 64 And Asc(Mid(mm, i, 1)) < 91) Or _
          (Asc(Mid(mm, i, 1)) > 96 And Asc(Mid(mm, i, 1)) < 123) Then
              Dim yl As dlm
              Open App.Path & "zcb.lsn" For Random As #1 Len = Len(yl)
                    For j = 1 To LOF(1) / Len(yl)
                        Get #1, j, yl
                        If yl.mz = dl.mz Then
                           yl.mm = mm
                           dl = yl
                           dr = True
                           Exit For
                        End If
                    Next j
                    Put #1, j, yl
                    MsgBox "�����޸ĳɹ���", 0, "��ʾ"
              Close #1
          End If
          If dr = True Then Exit For
      Next i
  End If
Loop Until dr = True
End Sub

Private Sub xz_Click(index As Integer)
Dim h1$, h2$, b$
Select Case index
        Case 1
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 + 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2 + 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 2
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2 + 1)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 3
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 4
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2 - 1)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 5
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2 - 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 6
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 7
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 - 1)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 8
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 - 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 9
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2, (dhz - 1) / 2 - 1)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 10
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2, (dhz - 1) / 2 - 2)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 11
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 - 1) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 - 1, (dhz - 1) / 2 - 1)
          b = bstr((dhz - 1) / 2 + 1, (dhz - 1) / 2 + 1)
        Case 12
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 - 1) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2 - 1, (dhz - 1) / 2 - 2)
          b = bstr((dhz - 1) / 2 + 2, (dhz - 1) / 2 + 1)
End Select
tcys.Enabled = False
ys.Enabled = False
Comhy.Visible = False
Comby.Visible = False
dzms.Enabled = False
qpsz.Enabled = False
szqp.Enabled = False
zykj.Enabled = False
If ys1 = 0 And ys2 = 0 Then
   ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
End If
Dim shu%, m1%, n1%, m2%, n2%, m3%, n3%
If md = 2 Then
   shu = MsgBox("��ѡ���ķ��������(Ҳ�����ߴ˾ְ���)��" & vbCrLf & "���ǡ�Ϊ���ԣ�����Ϊ���", 36, "��ѡ��")
   Picture1.Cls
   Call hqp
   If shu = vbYes Then
      hz(1) = b: bz(1) = h1: bz(2) = h2
      Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
      wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
      Call hqz(m1, n1, 1, False)
      Call hqz(m2, n2, 2, False)
      Call hqz(m3, n3, 1, True)
      Call jstr(autolz(wz()), m1, n1)
      Labelzb.Caption = Chr(65 + m1) & n1 + 1 & " " & "���ӵ㣺" & Chr(65 + m1) & n1 + 1
      lzd = "���ӵ㣺" & Chr(65 + m1) & n1 + 1
      Call hqz(m1, n1, 2, True)
      wz(m1, n1) = "����"
      bsh = 1
      bsh = bsh + 1
      hz(bsh) = bstr(m1, n1)
      bsb = 2
      Labelbsb.Caption = "��" & bsb & "��"
      Labelbsh.Caption = "��" & bsh & "��"
      slz = True
      If fzqz.Checked = False Then
         Call fk(slz)
      End If
      slz = False
      Timerb.Enabled = True
      Option2.Value = True
      Frame1.Caption = "��Ϸ�У�����ѡ��"
      Frame1.Enabled = False
      Randomize
      Dim rand%
      rand = Rnd * 3
      If rand < 1 Then
         rand = 2
      ElseIf rand < 2 Then
             rand = 3
      ElseIf rand < 3 Then
             rand = 6
      End If
      If rand = 2 Then
         Call ssjs_Click
         Call sijs_Click
      End If
      If rand = 3 Then
         Call cljs_Click
         Call ssjs_Click
      End If
      If rand = 6 Then
         Call sijs_Click
         Call cljs_Click
      End If
      Call qxzt
   ElseIf shu = vbNo Then
          hz(1) = h1: hz(2) = h2: bz(1) = b
          Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
          wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
          Call hqz(m1, n1, 1, True)
          Call hqz(m2, n2, 2, True)
          Call hqz(m3, n3, 1, False)
          slz = False
          bsh = 2: bsb = 1
          Labelbsh.Caption = "��" & bsh & "��"
          Labelbsb.Caption = "��" & bsb & "��"
          Timerb = True
          Option1.Value = True
          Frame1.Caption = "��Ϸ�У�����ѡ��"
          Frame1.Enabled = False
   End If
ElseIf md = 1 Then
       Picture1.Cls
       Call hqp
       hz(1) = h1: hz(2) = h2: bz(1) = b
       Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
       wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
       Call hqz(m1, n1, 1, True)
       Call hqz(m2, n2, 2, True)
       Call hqz(m3, n3, 1, False)
       slz = False
       bsh = 2: bsb = 1
       Labelbsh.Caption = "��" & bsh & "��"
       Labelbsb.Caption = "��" & bsb & "��"
       Timerb = True
       Option1.Value = True
       Frame1.Caption = "��Ϸ�У�����ѡ��"
       Frame1.Enabled = False
ElseIf md = 1 Then
       Picture1.Cls
       Call hqp
       hz(1) = h1: hz(2) = h2: bz(1) = b
       Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
       wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
       Call hqz(m1, n1, 1, True)
       Call hqz(m2, n2, 2, True)
       Call hqz(m3, n3, 1, False)
       slz = False
       bsh = 2: bsb = 1
       Labelbsh.Caption = "��" & bsh & "��"
       Labelbsb.Caption = "��" & bsb & "��"
       Timerb = True
       Option1.Value = True
       Frame1.Caption = "��Ϸ�У�����ѡ��"
       Frame1.Enabled = False
ElseIf md = 3 Then
       Dim zc$
       zc = Right(zhc, 3)
       If zc = "zhu" Then
          zhc = "zhu"
          shu = MsgBox("��ѡ���ķ��������(Ҳ�����ߴ˾ְ���)��" & vbCrLf & "���ǡ�Ϊ" & Lal1 & "������Ϊ" & Lal2, 36, "��ѡ��")
          Picture1.Cls
          Call hqp
          If shu = vbYes Then
             hz(1) = b: bz(1) = h1: bz(2) = h2
             Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
             wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
             Call hqz(m1, n1, 1, False)
             Call hqz(m2, n2, 2, False)
             Call hqz(m3, n3, 1, True)
             bsh = 1: bsb = 2
             Labelbsb.Caption = "��" & bsb & "��"
             Labelbsh.Caption = "��" & bsh & "��"
             slz = True
             If fzqz.Checked = False Then
                Call fk(Not slz)
             End If
             Timerh.Enabled = True
             Option2.Value = True
             Frame1.Caption = "��Ϸ�У�����ѡ��"
             Frame1.Enabled = False
             Call qxzt
             Picture2.Enabled = True
             If Win(1).State = sckConnected Then
                Win(1).SendData ("Y" & index & "xz")
             End If
          ElseIf shu = vbNo Then
                 hz(1) = h1: hz(2) = h2: bz(1) = b
                 Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                 wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                 Call hqz(m1, n1, 1, True)
                 Call hqz(m2, n2, 2, True)
                 Call hqz(m3, n3, 1, False)
                 slz = False
                 If fzqz.Checked = False Then
                    Call fk(Not slz)
                 End If
                 bsh = 2: bsb = 1
                 Labelbsh.Caption = "��" & bsh & "��"
                 Labelbsb.Caption = "��" & bsb & "��"
                 Timerb = True
                 Option1.Value = True
                 Frame1.Caption = "��Ϸ�У�����ѡ��"
                 Frame1.Enabled = False
                 Picture2.Enabled = False
                 If Win(1).State = sckConnected Then
                    Win(1).SendData ("N" & index & "xz")
                 End If
          End If
       ElseIf zc = "bei" Then
              zc = Left(zhc, 1)
              zhc = "bei"
              Picture1.Cls
              Call hqp
              If zc = "Y" Then
                 hz(1) = h1: hz(2) = h2: bz(1) = b
                 Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                 wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                 Call hqz(m1, n1, 1, True)
                 Call hqz(m2, n2, 2, True)
                 Call hqz(m3, n3, 1, False)
                 slz = False
                 If fzqz.Checked = False Then
                    Call fk(Not slz)
                 End If
                 bsh = 2: bsb = 1
                 Labelbsh.Caption = "��" & bsh & "��"
                 Labelbsb.Caption = "��" & bsb & "��"
                 Timerb = True
                 Option1.Value = True
                 Frame1.Caption = "��Ϸ�У�����ѡ��"
                 Frame1.Enabled = False
                 Picture2.Enabled = False
                 MsgBox Lal2 & "ʹ��ְҵ���֣�" & xz(index).Caption, 0, "��ʾ"
              ElseIf zc = "N" Then
                     hz(1) = b: bz(1) = h1: bz(2) = h2
                     Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                     wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                     Call hqz(m1, n1, 1, False)
                     Call hqz(m2, n2, 2, False)
                     Call hqz(m3, n3, 1, True)
                     bsh = 1: bsb = 2
                     Labelbsb.Caption = "��" & bsb & "��"
                     Labelbsh.Caption = "��" & bsh & "��"
                     slz = True
                     If fzqz.Checked = False Then
                        Call fk(Not slz)
                     End If
                     Timerh.Enabled = True
                     Option2.Value = True
                     Frame1.Caption = "��Ϸ�У�����ѡ��"
                     Frame1.Enabled = False
                     Call qxzt
                     Picture2.Enabled = True
                     MsgBox Lal2 & "ʹ��ְҵ���֣�" & xz(index).Caption, 0, "��ʾ"
              End If
       End If
End If
End Sub

Private Sub xztp_Click()
Picture1.Top = -15
Picture1.Left = -15
Picture1.Height = Picture2.Height
Picture1.Width = Picture2.Width
Picture1.Scale (0, dhz * 10)-(dhz * 10, 0)
Picture1 = LoadPicture("")
HS.Visible = False
VS.Visible = False
Call hqp
If fzqz.Checked = False Then
For i = 0 To 24
    For j = 0 To 24
        If wz(i, j) = "����" Then
           For m = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
           Next m
        End If
        If wz(i, j) = "����" Then
           For l = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
           Next l
        End If
    Next j
Next i
If bsb >= 1 Or bsh >= 1 Then
   If slz = True Then
      Call fk(Not slz)
   ElseIf slz = False Then
          Call fk(Not slz)
   End If
End If
ElseIf fzqz.Checked = True Then
       Dim qh%, qz%
       If bsh >= 1 Then
          For i = 1 To bsh
              Call jstr(hz(i), qh, qz)
              Imah(i).Left = qh * 10
              Imah(i).Top = (qz + 1) * 10
          Next i
       End If
       If bsb >= 1 Then
          For i = 1 To bsb
              Call jstr(bz(i), qh, qz)
              Imab(i).Left = qh * 10
              Imab(i).Top = (qz + 1) * 10
          Next i
       End If
End If
End Sub

Private Sub yanse_Click()
Call sdys_Click
End Sub

Private Sub Form_Load()  '������趨
res = False
Picture2.Height = Frame2.Top + Frame2.Height - Frame1.Top - HS.Height
Picture2.Width = Picture2.Height
VS.Left = Picture2.Left + Picture2.Width + 20
HS.Top = Picture2.Top + Picture2.Height + 20
Picsta.Top = HS.Top + HS.Height
Me.Height = Picsta.Top + Picsta.Height * 4.5
Me.Top = (Screen.Height - Me.Height) / 2
'//////////////////////////////////////////////////////////////
Picture1.Left = -15
Picture1.Top = -15
Picture1.Height = Picture2.Height
Picture1.Width = Picture2.Width
Picture1.Scale (0, 150)-(150, 0)                     '������
Pich.Height = Pich.Width
Picb.Height = Picb.Width
Pich.BackColor = Me.BackColor
Picb.BackColor = Me.BackColor
'//////////////////////////////////////////////////////////////
Pich.Scale (0, Pich.Height)-(Pich.Width, 0)
ys1 = RGB(0, 0, 0)
For i = 0 To Pich.Height / 2
    Pich.Circle (Pich.Width / 2, Pich.Width / 2), i, RGB(0, 0, 0)
Next i
Picb.Scale (0, Picb.Height)-(Picb.Width, 0)
ys2 = RGB(255, 255, 255)
For j = 0 To Picb.Height / 2
    Picb.Circle (Picb.Width / 2, Picb.Width / 2), j, RGB(255, 255, 255)
Next j
'//////////////////////////////////////////////////////////////
Timerh.Enabled = False: Timerb.Enabled = False
VS.Visible = False: HS.Visible = False
bt = True: bcts.Checked = True
qxts.Checked = True: dhz = 15
Option2.Value = True
pt1 = 150 / Picture2.Height: pt2 = 150 / Picture2.Width
jgx.Checked = True: jgx.Enabled = False
fsx.Checked = False: fsx.Enabled = True
ys.Enabled = False: tcys.Enabled = False
Comzdy.BackColor = RGB(195, 205, 99)
Comks(1).BackColor = RGB(195, 205, 99)
Comhq.BackColor = RGB(195, 205, 99)
Combc.BackColor = RGB(195, 205, 99)
Comhy.BackColor = RGB(195, 205, 99)
Comby.BackColor = RGB(195, 205, 99)
Comlj.BackColor = RGB(195, 205, 99)
Comjl.BackColor = RGB(195, 205, 99)
For i = 1 To 9
    If i = 4 Then
       qp(i).Checked = True
       qp(i).Enabled = False
       dhz = 15
       Me.Caption = "ɭ��������" & "��15��15��׼���̡�"
       pq(i).Checked = True
       pq(i).Enabled = False
    Else
        pq(i).Checked = False
        pq(i).Enabled = True
        qp(i).Checked = False
        qp(i).Enabled = True
    End If
    If i <= 8 Then
       shi(i).Checked = False
    End If
    If i <= 4 Then
       bu(i).Checked = False
    End If
Next i
For i = 5 To 16
    Load bu(i)
    If i <= 14 Then
       bu(i).Visible = False
    End If
Next i
bu(15).Caption = "�Զ��岽��"
bu(16).Caption = "ȡ �� �� ��"
Imah(1).Visible = False
Imab(1).Visible = False
ssjs.Checked = False: sijs.Checked = False
cljs.Checked = False
For i = 2 To 313
    Load Imah(i)
    Load Imab(i)
    Imah(i).Visible = False
    Imab(i).Visible = False
Next i
Call hqp
'//////////////////////////////////////////////////////////////
Lal1.Left = VS.Left + VS.Width + 100                 '���¡��ڷ��������λ��
Lal2.Left = Lal1.Left                                '���¡��׷��������λ��
Labelsjh.Left = Lal1.Left                            '��ʱ���ǩ��λ��
Labelsjb.Left = Lal2.Left
Labelbsh.Left = Lal1.Left + Lal1.Width               '��������ǩ��λ��
Labelbsb.Left = Lal2.Left + Lal2.Width
Labelbsh.Top = Lal1.Top                              '��������ǩ��λ��
Labelbsb.Top = Lal2.Top
Comhy.Left = Lal1.Left + Lal1.Width                  '��������ɫ��ťλ��
Comby.Left = Lal2.Left + Lal2.Width
Me.Width = Lal1.Left + Lal1.Width + Comhy.Width + 550
Me.Left = (Screen.Width - Me.Width) / 2
Comks(1).Left = Lal1.Left                               '����ʼ��ťλ��
Comks(1).Height = Comhq.Height
Comks(1).Top = Picture2.Top + Picture2.Height - Comks(1).Height * 2
Combc.Height = Comhq.Height                          '����������֡���ťλ��
Combc.Top = Comks(1).Top - Combc.Height - 100
Combc.Left = Lal1.Left
Comhq.Left = Lal1.Left                               '�������塱��ťλ��
Comhq.Top = Combc.Top - Comhq.Height - 100
Labelsj.Height = Picsta.Height                       '��ϵͳʱ��λ��
Labelsj.Top = 0
Labelsj.Left = Me.Width - Labelsj.Width - 300
Labelzb.Height = Picsta.Height                       '��ʵʱ����λ��
Labelzb.Top = 0
Labelzb.Left = Labelsj.Left - Labelzb.Width
Labeldlm.Height = Picsta.Height                      '����ʾ��¼��ҵ�λ��
Labeldlm.Top = 0
Labeldlm.Left = 0
Labeldlm.Caption = "��¼��ң�" & dl.mz
Picip.Left = Labeldlm.Left + Labeldlm.Width
Picip.Top = 0
Picip.Height = Picsta.Height
Labip.Left = 0
Labip.Top = 0
Labip.Height = Picip.Height
Labip.Width = Picip.Width
Labelts.Height = Picsta.Height                       '����ʾ���͵�λ��
Labelts.Top = 0
Labelts.Left = Picip.Left + Picip.Width
Picxz.Left = Labelts.Left + Labelts.Width
Picxz.Height = Picsta.Height
Picxz.Top = 0
Picxz.Width = Labelzb.Left - Labelts.Left - Labelts.Width
Labelxz.Top = 0: Labelxz.Left = 0
Labelxz.Height = Picxz.Height
Pich.Top = Lal1.Top + Lal1.Height                    '����ʾ������ɫ��ͼƬ��λ�ü���С
Pich.Left = Lal1.Left
Pich.Height = Pich.Width
Picb.Top = Lal2.Top + Lal2.Height
Picb.Left = Lal2.Left
Picb.Height = Picb.Width
Textdh.Top = Comhq.Top - Textdh.Height
Textdh.Left = Lal1.Left
Textdh.Width = Comby.Left + Comby.Width - Textdh.Left
lisdh.Top = Labelsjb.Top + Labelsjb.Height * 2
lisdh.Left = Lal1.Left
lisdh.Width = Textdh.Width
lisdh.Height = Textdh.Top - lisdh.Top
windowsh = Me.Height: windowsw = Me.Width
'//////////////////////////////////////////////////////////////
If md = 1 Then
   Call drdz_Click
ElseIf md = 2 Then
       Call rjdz_Click
ElseIf md = 3 Then
       Call wldz_Click
End If
res = True
End Sub

Private Sub Image1_Click(index As Integer)
Picture1.Top = -15
Picture1.Left = -15
Picture1.Height = Picture2.Height
Picture1.Width = Picture2.Width
Picture1.Scale (0, dhz * 10)-(dhz * 10, 0)
Picture1 = Image1(index).Picture
If Picture1.Width < Picture2.Width And Picture1.Height < Picture2.Height Then
   HS.Visible = False
   VS.Visible = False
   Picture1.Height = Picture2.Height
   Picture1.Width = Picture2.Width
End If
If Picture1.Height > Picture2.Height Then
   VS.Visible = True
   VS.Height = Picture2.Height
   If Picture1.Height - Picture2.Height > 32767 Then
      VS.Max = 32767
      v1 = (Picture1.Height - Picture2.Height) / 32767
   Else
       v1 = 0
       VS.Max = Picture1.Height - Picture2.Height
   End If
   VS.SmallChange = VS.Max / 20
   VS.LargeChange = VS.Max / 20
   VS.Value = 0
Else
   Picture1.Height = Picture2.Height
   VS.Visible = False
End If
If Picture1.Width > Picture2.Width Then
   HS.Visible = True
   HS.Width = Picture2.Width
   If Picture1.Width - Picture2.Width > 32767 Then
      HS.Max = 32767
      h1 = (Picture1.Width - Picture2.Width) / 32767
   Else
       h1 = 0
       HS.Max = Picture1.Width - Picture2.Width
   End If
   HS.SmallChange = HS.Max / 20
   HS.LargeChange = HS.Max / 20
   HS.Value = 0
Else
   Picture1.Width = Picture2.Width
   HS.Visible = False
End If
   Call hqp
   If fzqz.Checked = False Then
For i = 0 To 24
    For j = 0 To 24
        If wz(i, j) = "����" Then
           For m = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
           Next m
        End If
        If wz(i, j) = "����" Then
           For l = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
           Next l
        End If
    Next j
Next i
If bsb >= 1 Or bsh >= 1 Then
   If slz = True Then
      Call fk(Not slz)
   ElseIf slz = False Then
          Call fk(Not slz)
   End If
End If
End If
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim yj$
If Button = 1 And md = 1 Then
   If Frame1.Enabled = True Then  '�ж���Ϸ��ʼʱ���õ��ķ����߼�����
      Frame1.Caption = "��Ϸ��ʼ����ѡ��"
      If Option1.Value = True Then
         Frame1.Caption = "��Ϸ�У�����ѡ��"
         Frame1.Enabled = False
         slz = True
      ElseIf Option2.Value = True Then
         Frame1.Caption = "��Ϸ�У�����ѡ��"
         Frame1.Enabled = False
         slz = False
      End If
      tcys.Enabled = False
      ys.Enabled = False
      Comhy.Visible = False
      Comby.Visible = False
      dzms.Enabled = False
      qpsz.Enabled = False
      szqp.Enabled = False
      zykj.Enabled = False
      If ys1 = 0 And ys2 = 0 Then
         ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
      End If
   End If
   If Button = 1 And slz = True Then   '���ӡ���¼����λ�ú�������ɫ���ݡ��ж��Ƿ��������ߵȵȺܶ�
      Call pd(X, Y, hzb, zzb)
      If threej <> "" Then
         For i = 1 To (Len(threej) - 2) / 4
             If Mid(threej, 1 + (i - 1) * 4, 4) & Right(threej, 2) = bstr(hzb, zzb) & "�ڷ�" Then
                MsgBox "�˵��������֣��������ӣ�", 48, "��ʾ"
                Exit Sub
             End If
         Next i
      End If
      If fourj <> "" Then
             For i = 1 To (Len(fourj) - 2) / 4
                 If Mid(fourj, 1 + (i - 1) * 4, 4) & Right(fourj, 2) = bstr(hzb, zzb) & "�ڷ�" Then
                    MsgBox "�˵����Ľ��֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
             Next i
      End If
      If Option1.Value = True And cljs.Checked = True Then
             If jsix(hzb, zzb, slz) = True Then
                 MsgBox "�˵㳤�����֣��������ӣ�", 48, "��ʾ"
                 Exit Sub
             End If
      End If
      threej = "": fourj = ""
      If wz(hzb, zzb) = "" Then
         Labelzb.Caption = Chr(65 + hzb) & zzb + 1 & " " & "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
         lzd = "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
         Call hqz(hzb, zzb, bsh + 1, slz)
         slz = False
         wz(hzb, zzb) = "����"
         bsh = bsh + 1
         hz(bsh) = bstr(hzb, zzb)
         Labelbsh.Caption = "��" & bsh & "��"
         If fzqz.Checked = False Then
            Call fk(Not slz)
         End If
         tsbc = False
         Timerb.Enabled = True
         Timerh.Enabled = False
         Call qxzt
         If pdwz(wz(), hzb, zzb, "�ڷ�") = True Then
            Exit Sub
         End If
      End If
   ElseIf Button = 1 And slz = False Then
          Call pd(X, Y, hzb, zzb)
          If threej <> "" Then
             For i = 1 To (Len(threej) - 2) / 4
                 If Mid(threej, 1 + (i - 1) * 4, 4) & Right(threej, 2) = bstr(hzb, zzb) & "�׷�" Then
                    MsgBox "�˵��������֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
             Next i
          End If
          If fourj <> "" Then
                 For i = 1 To (Len(fourj) - 2) / 4
                     If Mid(fourj, 1 + (i - 1) * 4, 4) & Right(fourj, 2) = bstr(hzb, zzb) & "�׷�" Then
                        MsgBox "�˵����Ľ��֣��������ӣ�", 48, "��ʾ"
                        Exit Sub
                     End If
                 Next i
          End If
          If Option2.Value = True And cljs.Checked = True Then
                 If jsix(hzb, zzb, slz) = True Then
                    MsgBox "�˵㳤�����֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
          End If
          threej = "": fourj = ""
          If wz(hzb, zzb) = "����" Or wz(hzb, zzb) = "����" Then  '�жϸô��������ӣ������ô��Ƿ����
          ElseIf wz(hzb, zzb) = "" Then
                 Labelzb.Caption = Chr(65 + hzb) & zzb + 1 & " " & "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
                 lzd = "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
                 Call hqz(hzb, zzb, bsb + 1, slz)
                 slz = True
                 wz(hzb, zzb) = "����"
                 bsb = bsb + 1
                 bz(bsb) = bstr(hzb, zzb)
                 Labelbsb.Caption = "��" & bsb & "��"
                 If fzqz.Checked = False Then
                    Call fk(Not slz)
                 End If
                 tsbc = False
                 Timerb.Enabled = False
                 Timerh.Enabled = True
                 Call qxzt
                 If pdwz(wz(), hzb, zzb, "�׷�") = True Then
                    Exit Sub
                 End If
          End If
   End If
ElseIf Button = 1 And md = 2 Then
       Dim ah%, az%
       If Frame1.Enabled = True Then  '�ж���Ϸ��ʼʱ���õ��ķ����߼�����
          Frame1.Caption = "��Ϸ��ʼ����ѡ��"
          If Option1.Value = True Then
             Frame1.Caption = "��Ϸ�У�����ѡ��"
             Frame1.Enabled = False
             slz = True
          ElseIf Option2.Value = True Then
             Frame1.Caption = "��Ϸ�У�����ѡ��"
             Frame1.Enabled = False
             slz = False
          End If
          tcys.Enabled = False
          ys.Enabled = False
          Comhy.Visible = False
          Comby.Visible = False
          dzms.Enabled = False
          qpsz.Enabled = False
          szqp.Enabled = False
          zykj.Enabled = False
          If ys1 = 0 And ys2 = 0 Then
             ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
          End If
          Randomize
          Dim rand%
          rand = Rnd * 3
          If rand < 1 Then
             rand = 2
          ElseIf rand < 2 Then
                 rand = 3
          ElseIf rand < 3 Then
                 rand = 6
          End If
          If rand = 2 Then
             Call ssjs_Click
             Call sijs_Click
          End If
          If rand = 3 Then
             Call cljs_Click
             Call ssjs_Click
         End If
         If rand = 6 Then
            Call sijs_Click
            Call cljs_Click
         End If
       End If
       If slz = False Then
          Call pd(X, Y, hzb, zzb)
          If threej <> "" Then
             For i = 1 To (Len(threej) - 4) / 4
                 If Mid(threej, 1 + (i - 1) * 4, 4) & Right(threej, 4) = bstr(hzb, zzb) & dl.mz Then
                    MsgBox "�˵��������֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
             Next i
          End If
          If fourj <> "" Then
                 For i = 1 To (Len(fourj) - 4) / 4
                     If Mid(fourj, 1 + (i - 1) * 4, 4) & Right(fourj, 4) = bstr(hzb, zzb) & dl.mz Then
                        MsgBox "�˵����Ľ��֣��������ӣ�", 48, "��ʾ"
                        Exit Sub
                     End If
                 Next i
          End If
          If Option2.Value = True And cljs.Checked = True Then
                 If jsix(hzb, zzb, slz) = True Then
                    MsgBox "�˵㳤�����֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
          End If
          threej = "": fourj = ""
          If wz(hzb, zzb) = "" Then
             Labelzb.Caption = Chr(65 + hzb) & zzb + 1 & " " & "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
             lzd = "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
             wz(hzb, zzb) = "����"
             Call hqz(hzb, zzb, bsb + 1, slz)
             bsb = bsb + 1
             bz(bsb) = bstr(hzb, zzb)
             Labelbsb.Caption = "��" & bsb & "��"
             If fzqz.Checked = False Then
                Call fk(slz)
             End If
             tsbc = False
             Timerb.Enabled = False
             slz = True
             If pdwz(wz(), hzb, zzb, dl.mz) = True Then
                Exit Sub
             End If
             '/////////////////////����Ϊ��ң�����Ϊ����
             Dim ran%, rh%, rz%, ih%, iz%
             Labelsjh.Caption = "����˼���С���"
             If bsb = 1 And bsh = 0 Then   '�������һ��
                Do
                  Do
                   Randomize
                   ran = Int(Rnd * 8)
                   Select Case ran
                         Case 0
                              rh = hzb: rz = zzb - 1
                         Case 1
                              rh = hzb - 1: rz = zzb + 1
                         Case 2
                              rh = hzb + 1: rz = zzb
                         Case 3
                              rh = hzb: rz = zzb + 1
                         Case 4
                              rh = hzb + 1: rz = zzb + 1
                         Case 5
                              rh = hzb + 1: rz = zzb - 1
                         Case 6
                              rh = hzb - 1: rz = zzb - 1
                         Case 7
                              rh = hzb - 1: rz = zzb
                   End Select
                  Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                Loop Until wz(rh, rz) = ""
                Labelzb.Caption = Chr(65 + rh) & rz + 1 & " " & "���ӵ㣺" & Chr(65 + rh) & rz + 1
                lzd = "���ӵ㣺" & Chr(65 + rh) & rz + 1
                Call hqz(rh, rz, bsh + 1, slz)
                wz(rh, rz) = "����"
                bsh = bsh + 1
                hz(bsh) = bstr(rh, rz)
                Labelbsh.Caption = "��" & bsh & "��"
                If fzqz.Checked = False Then
                   Call fk(slz)
                End If
                slz = False
                Timerb.Enabled = True
                Labelsjh.Caption = ""
             ElseIf bsb = 1 And bsh = 1 Then    '��������һ��
                Call jstr(hz(1), ih, iz)
                If zzb = iz Then
                   Do
                     Do
                       Randomize
                       ran = Int(Rnd * 6)
                       Select Case ran
                         Case 0
                              rh = ih + 1: rz = iz - 1
                         Case 1
                              rh = ih - 1: rz = iz - 1
                         Case 2
                              rh = ih: rz = iz + 1
                         Case 3
                              rh = ih + 1: rz = iz + 1
                         Case 4
                              rh = ih - 1: rz = iz + 1
                         Case 5
                              rh = ih: rz = iz - 1
                       End Select
                     Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                   Loop Until wz(rh, rz) = ""
                ElseIf hzb = ih Then
                       Do
                         Do
                           Randomize
                           ran = Int(Rnd * 6)
                           Select Case ran
                                 Case 0
                                      rh = ih - 1: rz = iz - 1
                                 Case 1
                                      rh = ih + 1: rz = iz + 1
                                 Case 2
                                      rh = ih - 1: rz = iz
                                 Case 3
                                      rh = ih - 1: rz = iz + 1
                                 Case 4
                                      rh = ih - 1: rz = iz
                                 Case 5
                                      rh = ih + 1: rz = iz - 1
                           End Select
                         Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                       Loop Until wz(rh, rz) = ""
                ElseIf hzb <> ih And zzb <> iz Then
                       Do
                         Do
                           Randomize
                           ran = Int(Rnd * 8)
                           Select Case ran
                                 Case 0
                                      rh = ih: rz = iz - 1
                                 Case 1
                                      rh = ih + 1: rz = iz
                                 Case 2
                                      rh = ih - 1: rz = iz + 1
                                 Case 3
                                      rh = ih - 1: rz = iz
                                 Case 4
                                      rh = ih + 1: rz = iz - 1
                                 Case 5
                                      rh = ih: rz = iz + 1
                                 Case 6
                                      rh = ih + 1: rz = iz + 1
                                 Case 7
                                      rh = ih - 1: rz = iz - 1
                           End Select
                         Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                       Loop Until (wz(rh, rz) = "" And Abs(hzb - rh) <> 2) Or (wz(rh, rz) = "" And Abs(zzb - rz) <> 2)
                End If
                Labelzb.Caption = Chr(65 + rh) & rz + 1 & " " & "���ӵ㣺" & Chr(65 + rh) & rz + 1
                lzd = "���ӵ㣺" & Chr(65 + rh) & rz + 1
                Call hqz(rh, rz, bsh + 1, slz)
                wz(rh, rz) = "����"
                bsh = bsh + 1
                hz(bsh) = bstr(rh, rz)
                Labelbsh.Caption = "��" & bsh & "��"
                If fzqz.Checked = False Then
                   Call fk(slz)
                End If
                slz = False
                Timerb.Enabled = True
                Labelsjh.Caption = ""
                Call qxzt
             ElseIf bsb >= 2 And bsb <= 3 And (bsh = 1 Or bsh = 2) Then '�������
                Dim h2%, z2%
                If er(rh, rz, Not slz) = True Then
                Else
                    If bsh = 1 Then
                       Call jstr(hz(1), h2, z2)
                           Do
                             Do
                               Randomize
                               ran = Int(Rnd * 8)
                               Select Case ran
                                    Case 0
                                         rh = h2: rz = z2 - 1
                                    Case 1
                                         rh = h2 - 1: rz = z2 + 1
                                    Case 2
                                         rh = h2 + 1: rz = z2
                                    Case 3
                                         rh = h2: rz = z2 + 1
                                    Case 4
                                         rh = h2 + 1: rz = z2 + 1
                                    Case 5
                                         rh = h2 + 1: rz = z2 - 1
                                    Case 6
                                         rh = h2 - 1: rz = z2 - 1
                                    Case 7
                                         rh = h2 - 1: rz = z2
                               End Select
                             Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                           Loop Until wz(rh, rz) = ""
                    ElseIf bsh = 2 And er(rh, rz, slz) = True Then
                    ElseIf bsh = 2 And er(rh, rz, slz) = False Then
                           Call jstr(hz(2), h2, z2)
                           Do
                             Do
                               Randomize
                               ran = Int(Rnd * 8)
                               Select Case ran
                                    Case 0
                                         rh = h2: rz = z2 - 1
                                    Case 1
                                         rh = h2 - 1: rz = z2 + 1
                                    Case 2
                                         rh = h2 + 1: rz = z2
                                    Case 3
                                         rh = h2: rz = z2 + 1
                                    Case 4
                                         rh = h2 + 1: rz = z2 + 1
                                    Case 5
                                         rh = h2 + 1: rz = z2 - 1
                                    Case 6
                                         rh = h2 - 1: rz = z2 - 1
                                    Case 7
                                         rh = h2 - 1: rz = z2
                               End Select
                             Loop Until rh >= 0 And rh <= dhz - 1 And rz >= 0 And rz <= dhz - 1
                           Loop Until wz(rh, rz) = ""
                    End If
                End If
                Labelzb.Caption = Chr(65 + rh) & rz + 1 & " " & "���ӵ㣺" & Chr(65 + rh) & rz + 1
                lzd = "���ӵ㣺" & Chr(65 + rh) & rz + 1
                Call hqz(rh, rz, bsh + 1, slz)
                wz(rh, rz) = "����"
                bsh = bsh + 1
                hz(bsh) = bstr(rh, rz)
                Labelbsh.Caption = "��" & bsh & "��"
                If fzqz.Checked = False Then
                   Call fk(slz)
                End If
                slz = False
                Timerb.Enabled = True
                Labelsjh.Caption = ""
                Call qxzt
             Else
                 Call jstr(autolz(wz()), ah, az)
                 Labelzb.Caption = Chr(65 + ah) & az + 1 & " " & "���ӵ㣺" & Chr(65 + ah) & az + 1
                 lzd = "���ӵ㣺" & Chr(65 + ah) & az + 1
                 Call hqz(ah, az, bsh + 1, slz)
                 wz(ah, az) = "����"
                 bsh = bsh + 1
                 hz(bsh) = bstr(ah, az)
                 Labelbsh.Caption = "��" & bsh & "��"
                 If fzqz.Checked = False Then
                    Call fk(slz)
                 End If
                 slz = False
                 Timerb.Enabled = True
                 Labelsjh.Caption = ""
                 Call qxzt
                 If pdwz(wz(), ah, az, "����") = True Then
                    Exit Sub
                 End If
             End If
          End If
       End If
ElseIf Button = 1 And md = 3 Then
       If slz = True And Button = 1 Then
          tcys.Enabled = False
          ys.Enabled = False
          Comhy.Visible = False
          Comby.Visible = False
          dzms.Enabled = False
          qpsz.Enabled = False
          szqp.Enabled = False
          zykj.Enabled = False
          tsbc = False
          If ys1 = 0 And ys2 = 0 Then
             ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
          End If
          Call pd(X, Y, hzb, zzb)
          If threej <> "" Then
             For i = 1 To (Len(threej) - 4) / 4
                 If Mid(threej, 1 + (i - 1) * 4, 4) & Right(threej, 4) = bstr(hzb, zzb) & Lal1 Then
                    MsgBox "�˵��������֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
             Next i
          End If
          If fourj <> "" Then
                 For i = 1 To (Len(fourj) - 4) / 4
                     If Mid(fourj, 1 + (i - 1) * 4, 4) & Right(fourj, 4) = bstr(hzb, zzb) & Lal1 Then
                        MsgBox "�˵����Ľ��֣��������ӣ�", 48, "��ʾ"
                        Exit Sub
                     End If
                 Next i
          End If
          If zhc = "zhu" And cljs.Checked = True Then
                 If jsix(hzb, zzb, slz) = True Then
                    MsgBox "�˵㳤�����֣��������ӣ�", 48, "��ʾ"
                    Exit Sub
                 End If
          End If
          threej = "": fourj = ""
          If wz(hzb, zzb) = "" Then
                Labelzb.Caption = Chr(65 + hzb) & zzb + 1 & " " & "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
                lzd = "���ӵ㣺" & Chr(65 + hzb) & zzb + 1
                Call hqz(hzb, zzb, bsh + 1, slz)
                slz = False
                wz(hzb, zzb) = "����"
                bsh = bsh + 1
                hz(bsh) = bstr(hzb, zzb)
                If Win(1).State = sckConnected Then
                   Win(1).SendData (hz(bsh) & "wz")
                End If
                Labelbsh.Caption = "��" & bsh & "��"
                If fzqz.Checked = False Then
                   Call fk(Not slz)
                End If
                Timerb.Enabled = True
                Timerh.Enabled = False
                Picture2.Enabled = False
                Call qxzt
                If pdwz(wz(), hzb, zzb, Lal1.Caption) = True Then
                   Exit Sub
                End If
          End If
       End If
End If
If Button = 2 Then
   If Picture1 = LoadPicture("") Then
      xztp.Enabled = False
   Else
       xztp.Enabled = True
   End If
   PopupMenu tc1, 0
End If
End Sub

Public Sub pd(ByVal X!, ByVal Y!, ByRef ch%, ByRef cz%)
If X < 10 And X >= 0 Then         '��������ͷλ���жϸô���������������Ϊ����
   ch = 0
ElseIf X < 100 Then
   ch = Val(Left(Trim(Str(X)), 1))
ElseIf X <= 250 Then
   ch = Val(Left(Trim(Str(X)), 2))
End If
If Y < 10 And Y >= 0 Then
   cz = 0
ElseIf Y < 100 Then
   cz = Val(Left(Trim(Str(Y)), 1))
ElseIf Y <= 250 Then
   cz = Val(Left(Trim(Str(Y)), 2))
End If
End Sub

Public Function pdwz(wz$(), ByVal ch%, ByVal cz%, ByVal nfy$) As Boolean
Dim jls%(1 To 20), yj$ '�Ƚϵļ�¼��
pdwz = False
If bsh = (dhz ^ 2 + 1) / 2 Or bsb = (dhz ^ 2 + 1) / 2 Then
   MsgBox "�˾ֺ��壡", 48, "���Ľ��"
   Dim yl As dlm, i%
   Open App.Path & "zcb.lsn" For Random As #1 Len = Len(yl)
        For i = 1 To LOF(1) / Len(yl)
            Get #1, i, yl
            If yl.mz = dl.mz Then
               Exit For
            End If
        Next i
        If md = 1 Then
           yl.drh.bs_t = bsh + yl.drh.bs_t
           yl.drh.sj_t = sjh + yl.drh.sj_t
           yl.drb.bs_t = bsb + yl.drb.bs_t
           yl.drb.sj_t = sjb + yl.drb.sj_t
           yl.drh.tie = yl.drh.tie + 1
           yl.drb.tie = yl.drb.tie + 1
        ElseIf md = 2 Then
               yl.rj.bs_t = bsb + yl.rj.bs_t
               yl.rj.sj_t = sjb + yl.rj.sj_t
               yl.rj.tie = yl.rj.tie + 1
        ElseIf md = 3 Then
               yl.wl.bs_t = bsh + yl.wl.bs_t
               yl.wl.sj_t = sjh + yl.wl.sj_t
               yl.wl.tie = yl.wl.tie + 1
        End If
        Put #1, i, yl
        dl = yl
   Close #1
   pdwz = True
   Exit Function
End If
For i = 1 To 4
    If cz <= dhz - 5 Then '1-3Ϊ���ϣ����Һ������ϵ���űȽ�
       If wz(ch, cz) = wz(ch, cz + i) And wz(ch, cz + i) <> "" Then
          jls(1) = jls(1) + 1
       End If
    End If
    If ch <= dhz - 5 Then
       If wz(ch, cz) = wz(ch + i, cz) And wz(ch + i, cz) <> "" Then
          jls(2) = jls(2) + 1
       End If
    End If
    If ch <= dhz - 5 And cz <= dhz - 5 Then
       If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
          jls(3) = jls(3) + 1
       End If
    End If
    If cz >= 4 Then  '4-6Ϊ���£�����������µ���űȽ�
       If wz(ch, cz) = wz(ch, cz - i) And wz(ch, cz - i) <> "" Then
          jls(4) = jls(4) + 1
       End If
    End If
    If ch >= 4 Then
       If wz(ch, cz) = wz(ch - i, cz) And wz(ch - i, cz) <> "" Then
          jls(5) = jls(5) + 1
       End If
    End If
    If ch >= 4 And cz >= 4 Then
       If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
          jls(6) = jls(6) + 1
       End If
    End If
    If ch >= 4 And cz <= dhz - 5 Then '7-8Ϊ��б�Ϻ���б�µ���űȽ�
       If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
          jls(7) = jls(7) + 1
       End If
    End If
    If ch <= dhz - 5 And cz >= 5 Then
       If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
          jls(8) = jls(8) + 1
        End If
    End If
    If cz >= 2 And cz <= dhz - 3 Then '9-12Ϊ���򣬺���,б�º�б�ϵ�ǰ�����űȽ�
       If wz(ch, cz - 2) = wz(ch, cz - 2 + i) And wz(ch, cz - 2 + i) <> "" Then
          jls(9) = jls(9) + 1
       End If
    End If
    If ch >= 2 And ch <= dhz - 3 Then
       If wz(ch - 2, cz) = wz(ch - 2 + i, cz) And wz(ch - 2, cz) <> "" Then
          jls(10) = jls(10) + 1
       End If
    End If
    If ch >= 2 And cz >= 2 And ch <= dhz - 3 And cz <= dhz - 3 Then
       If wz(ch - 2, cz - 2) = wz(ch - 2 + i, cz - 2 + i) And wz(ch - 2 + i, cz - 2 + i) <> "" Then
          jls(11) = jls(11) + 1
       End If
    End If
    If ch >= 2 And cz <= dhz - 3 And ch <= dhz - 3 And cz >= 2 Then
       If wz(ch - 2, cz + 2) = wz(ch - 2 + i, cz + 2 - i) And wz(ch - 2 + i, cz + 2 - i) <> "" Then
          jls(12) = jls(12) + 1
       End If
    End If
    If ch >= 1 And ch <= dhz - 4 Then     '13-17Ϊ���ң�����б�ϣ����ϣ�����б�ϣ������һ���űȽ�
       If wz(ch - 1, cz) = wz(ch - 1 + i, cz) And wz(ch - 1 + i, cz) <> "" Then
          jls(13) = jls(13) + 1
       End If
    End If
    If ch >= 1 And cz >= 1 And ch <= dhz - 4 And cz <= dhz - 4 Then
       If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
          jls(14) = jls(14) + 1
       End If
    End If
    If cz >= 1 And cz <= dhz - 4 Then
       If wz(ch, cz - 1) = wz(ch, cz - 1 + i) And wz(ch, cz - 1 + i) <> "" Then
          jls(15) = jls(15) + 1
       End If
    End If
    If ch <= dhz - 2 And cz >= 1 And ch >= 3 And cz <= dhz - 4 Then
       If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
          jls(16) = jls(16) + 1
       End If
    End If
    If ch >= 3 And ch <= dhz - 2 Then
       If wz(ch + 1, cz) = wz(ch + 1 - i, cz) And wz(ch + 1 - i, cz) <> "" Then
          jls(17) = jls(17) + 1
       End If
    End If
    If ch >= 3 And cz >= 3 And ch <= dhz - 2 And cz <= dhz - 2 Then '18-20Ϊ����б�£����£�����б�µ�һ���Ƚ�
       If wz(ch + 1, cz + 1) = wz(ch + 1 - i, cz + 1 - i) And wz(ch + 1 - i, cz + 1 - i) <> "" Then
          jls(18) = jls(18) + 1
       End If
    End If
    If cz >= 3 And cz <= dhz - 2 Then
       If wz(ch, cz + 1) = wz(ch, cz + 1 - i) And wz(ch, cz + 1 - i) <> "" Then
          jls(19) = jls(19) + 1
       End If
    End If
    If ch >= 1 And cz <= dhz - 2 And ch <= dhz - 4 And cz >= 3 Then
       If wz(ch - 1, cz + 1) = wz(ch - 1 + i, cz + 1 - i) And wz(ch - 1 + i, cz + 1 - i) <> "" Then
          jls(20) = jls(20) + 1
       End If
    End If
Next i
For j = 1 To 20
    If jls(j) = 4 Then
       MsgBox nfy & "��ʤ!", 48, "���Ľ��"
       If md = 1 Then
          If wz(ch, cz) = "����" Then
             yj = "�ڷ���ʤ"
          ElseIf wz(ch, cz) = "����" Then
                 yj = "�׷���ʤ"
          End If
       ElseIf md = 2 Then
              If wz(ch, cz) = "����" Then
                 yj = "���Ի�ʤ"
              ElseIf wz(ch, cz) = "����" Then
                 yj = dl.mz & "��ʤ"
              End If
       ElseIf md = 3 Then
              If wz(ch, cz) = "����" Then
                 yj = dl.mz & "��ʤ"
              ElseIf wz(ch, cz) = "����" Then
                 yj = Lal2.Caption & "��ʤ"
              End If
       End If
       Call sl(yj)
       pdwz = True
       Exit Function
    End If
Next j
If xzb > 0 Then
   Dim ad$
   If bsh >= xzb Then
      MsgBox "���Ʋ����ѵ���" & Lal2 & "��ʤ��", 48, "��ʾ"
      Call sl(ad)
      Exit Function
   End If
   If bsb >= xzb Then
      MsgBox "���Ʋ����ѵ���" & Lal1 & "��ʤ��", 48, "��ʾ"
      Call sl(ad)
      Exit Function
   End If
End If
End Function

Private Sub Timerb_Timer()
sjb = sjb + 1
Labelsjb.Caption = Lal2.Caption & "����ʱ" & sjb \ 60 & "��" & sjb - (sjb \ 60) * 60 & "��"
If xzs > 0 Then
   If sjb >= xzs Then
      Dim ad$
      MsgBox "����ʱ���ѵ���" & Lal1 & "��ʤ��", 48, "��ʾ"
      Call sl(ad)
   End If
End If
End Sub

Private Sub Timerh_Timer()
sjh = sjh + 1
Labelsjh.Caption = Lal1.Caption & "����ʱ" & sjh \ 60 & "��" & sjh - (sjh \ 60) * 60 & "��"
If xzs > 0 Then
   If sjh >= xzs Then
      Dim ad$
      MsgBox "����ʱ���ѵ���" & Lal2 & "��ʤ��", 48, "��ʾ"
      Call sl(ad)
   End If
End If
End Sub
Private Sub hfys_Click()  '�ڷ�������ɫ�趨
com.CancelError = True
On Error GoTo errhandler
com.DialogTitle = Lal1 & "������ɫѡȡ"
com.ShowColor
ys1 = com.Color
Pich.Scale (0, 10)-(10, 0)
Pich.Cls
Pich.BackColor = Me.BackColor
For i = 0 To 100
    Pich.Circle (5, 5), i / 20, ys1
Next i
If md = 3 Then
   If Win(1).State = sckConnected Then
      Win(1).SendData (ys1 & "ys")
   End If
End If
errhandler:
End Sub

Private Sub bfys_Click()   '�׷�������ɫ�趨
com.CancelError = True
On Error GoTo errhandler
com.DialogTitle = Lal2 & "������ɫѡȡ"
com.ShowColor
ys2 = com.Color
Picb.Cls
Picb.BackColor = Me.BackColor
Picb.Scale (0, 100)-(100, 0)
For j = 1 To 50
    Picb.Circle (50, 50), j, ys2
Next j
errhandler:
End Sub

Public Function three(wz$(), ByRef h%, ByRef z%, ByVal bo As Boolean) As Boolean  '�ж����͵��������(������������������ͷ��������ͷ��������һ������λ�ã�
Dim jls%(1 To 48), ch%, cz%, et$                 '�ú���Ϊ�棬˵��������Ϊ������
three = False                                       '������������ͷ������һ����λ��������λ��������һ����λ��
If bo = True Then                                'Ҳ������������أ��ͻ���������͡�
   et = "����"
Else
   et = "����"
End If
For ch = 0 To dhz - 1
  For cz = 0 To dhz - 1
    For k = 1 To 48
        jls(k) = 0
    Next k
    If wz(ch, cz) = et Then
    For i = 1 To 2
      If ch <= dhz - 5 And ch >= 2 Then    '�ң��ϣ����±Ƚ�
         If wz(ch, cz) = wz(ch + i, cz) And wz(ch + i, cz) <> "" Then
            If wz(ch - 1, cz) = "" And wz(ch + 3, cz) = "" Then
             If wz(ch - 2, cz) = "" Or wz(ch + 4, cz) = "" Then
                jls(1) = jls(1) + 1
             End If
          End If
       End If
    End If
    If cz <= dhz - 5 And cz >= 2 Then
       If wz(ch, cz) = wz(ch, cz + i) And wz(ch, cz + i) <> "" Then
          If wz(ch, cz - 1) = "" And wz(ch, cz + 3) = "" Then
             If wz(ch, cz - 2) = "" Or wz(ch, cz + 4) = "" Then
                jls(2) = jls(2) + 1
             End If
          End If
       End If
    End If
    If ch >= 4 And ch <= dhz - 3 Then
       If wz(ch, cz) = wz(ch - i, cz) And wz(ch - i, cz) <> "" Then
          If wz(ch + 1, cz) = "" And wz(ch - 3, cz) = "" Then
             If wz(ch + 2, cz) = "" Or wz(ch - 4, cz) = "" Then
                jls(3) = jls(3) + 1
             End If
          End If
       End If
    End If
    If cz >= 4 And cz <= dhz - 3 Then
       If wz(ch, cz) = wz(ch, cz - i) And wz(ch, cz - i) <> "" Then
          If wz(ch, cz + 1) = "" And wz(ch, cz - 3) = "" Then
             If wz(ch, cz + 2) = "" Or wz(ch, cz - 4) = "" Then
                jls(4) = jls(4) + 1
             End If
          End If
       End If
    End If
    If ch <= dhz - 5 And cz <= dhz - 5 And ch >= 2 And cz >= 2 Then '����б������б������б������б�Ƚ�
       If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
          If wz(ch - 1, cz - 1) = "" And wz(ch + 3, cz + 3) = "" Then
             If wz(ch - 2, cz - 2) = "" Or wz(ch + 4, cz + 4) = "" Then
                jls(5) = jls(5) + 1
             End If
          End If
       End If
    End If
    If ch >= 4 And cz <= dhz - 5 And ch <= dhz - 3 And cz >= 2 Then
       If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
          If wz(ch + 1, cz - 1) = "" And wz(ch - 3, cz + 3) = "" Then
             If wz(ch + 2, cz - 2) = "" Or wz(ch - 4, cz + 4) = "" Then
                jls(6) = jls(6) + 1
             End If
          End If
       End If
    End If
    If ch >= 4 And cz >= 4 And ch <= dhz - 3 And cz <= dhz - 3 Then
       If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
          If wz(ch + 1, cz + 1) = "" And wz(ch - 3, cz - 3) = "" Then
             If wz(ch + 2, cz + 2) = "" Or wz(ch - 4, cz - 4) = "" Then
                jls(7) = jls(7) + 1
             End If
          End If
       End If
    End If
    If ch <= dhz - 5 And cz >= 4 And ch >= 2 And cz <= dhz - 3 Then
       If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
          If wz(ch - 1, cz + 1) = "" And wz(ch + 3, cz - 3) = "" Then
             If wz(ch - 2, cz + 2) = "" Or wz(ch + 4, cz - 4) = "" Then
                jls(8) = jls(8) + 1
             End If
          End If
       End If
    End If
    If ch >= 3 And cz >= 3 And cz <= dhz - 4 And ch <= dhz - 4 Then '��б����б���ᣬ�ݱȽ�
       If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
          If wz(ch - 2, cz - 2) = "" And wz(ch + 2, cz + 2) = "" Then
             If wz(ch - 3, cz - 3) = "" Or wz(ch + 3, cz + 3) = "" Then
                jls(9) = jls(9) + 1
             End If
          End If
       End If
    End If
    If ch <= dhz - 4 And cz >= 3 And ch >= 3 And cz <= dhz - 4 Then
       If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
          If wz(ch + 2, cz - 2) = "" And wz(ch - 2, cz + 2) = "" Then
             If wz(ch + 3, cz - 3) = "" Or wz(ch - 3, cz + 3) = "" Then
                jls(10) = jls(10) + 1
             End If
          End If
       End If
    End If
    If ch >= 3 And ch <= dhz - 4 Then
       If wz(ch - 1, cz) = wz(ch - 1 + i, cz) And wz(ch - 1 + i, cz) <> "" Then
          If wz(ch - 2, cz) = "" And wz(ch + 2, cz) = "" Then
             If wz(ch - 3, cz) = "" Or wz(ch + 3, cz) = "" Then
                jls(11) = jls(11) + 1
             End If
          End If
       End If
    End If
    If cz >= 3 And cz <= dhz - 4 Then
       If wz(ch, cz - 1) = wz(ch, cz - 1 + i) And wz(ch, cz - 1 + i) <> "" Then
          If wz(ch, cz - 2) = "" And wz(ch, cz + 2) = "" Then
             If wz(ch, cz - 3) = "" Or wz(ch, cz + 3) = "" Then
                jls(12) = jls(12) + 1
             End If
          End If
       End If
    End If
    If ch = dhz - 4 Then
       If wz(ch, cz) = wz(ch + i, cz) And wz(ch + i, cz) <> "" Then
          If wz(ch - 1, cz) = "" And wz(ch + 3, cz) = "" Then
             If wz(ch - 2, cz) = "" Then
                jls(13) = jls(13) + 1
             End If
          End If
       End If
       If cz >= 2 And cz <= dhz - 4 Then
          If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
             If wz(ch - 1, cz - 1) = "" And wz(ch + 3, cz + 3) = "" Then
                If wz(ch - 2, cz - 2) = "" Then
                   jls(41) = jls(41) + 1
                End If
             End If
          End If
       End If
       If cz >= 3 And cz <= dhz - 3 Then
          If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
             If wz(ch - 1, cz + 1) = "" And wz(ch + 3, cz - 3) = "" Then
                If wz(ch - 2, cz + 2) = "" Then
                   jls(42) = jls(42) + 1
                End If
             End If
          End If
       End If
    End If
    If ch = 3 Then
       If wz(ch, cz) = wz(ch - i, cz) And wz(ch - i, cz) <> "" Then
          If wz(ch + 1, cz) = "" And wz(ch - 3, cz) = "" Then
             If wz(ch + 2, cz) = "" Then
                jls(14) = jls(14) + 1
             End If
          End If
       End If
       If cz >= 2 And cz <= dhz - 4 Then
          If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
             If wz(ch + 1, cz - 1) = "" And wz(ch - 3, cz + 3) = "" Then
                If wz(ch + 2, cz - 2) = "" Then
                   jls(43) = jls(43) + 1
                End If
             End If
          End If
       End If
       If cz >= 3 And cz <= dhz - 3 Then
          If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
             If wz(ch + 1, cz + 1) = "" And wz(ch - 3, cz - 3) = "" Then
                If wz(ch + 2, cz + 2) = "" Then
                   jls(44) = jls(44) + 1
                End If
             End If
          End If
       End If
    End If
    If cz = 3 Then
       If wz(ch, cz) = wz(ch, cz - i) And wz(ch, cz - i) <> "" Then
          If wz(ch, cz + 1) = "" And wz(ch, cz - 3) = "" Then
             If wz(ch, cz + 2) = "" Then
                jls(15) = jls(15) + 1
             End If
          End If
       End If
       If ch <= dhz - 3 And ch >= 3 Then
          If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
             If wz(ch + 1, cz + 1) = "" And wz(ch - 3, cz - 3) = "" Then
                If wz(ch + 2, cz + 2) = "" Then
                   jls(45) = jls(45) + 1
                End If
             End If
          End If
        End If
        If ch >= 2 And ch <= dhz - 4 Then
           If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
             If wz(ch - 1, cz + 1) = "" And wz(ch + 3, cz - 3) = "" Then
                If wz(ch - 2, cz + 2) = "" Then
                   jls(46) = jls(46) + 1
                End If
             End If
          End If
        End If
    End If
    If cz = dhz - 4 Then
       If wz(ch, cz) = wz(ch, cz + i) And wz(ch, cz + i) <> "" Then
          If wz(ch, cz - 1) = "" And wz(ch, cz + 3) = "" Then
             If wz(ch, cz - 2) = "" Then
                jls(16) = jls(16) + 1
             End If
          End If
       End If
       If ch >= 2 And ch <= dhz - 4 Then
          If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
             If wz(ch - 1, cz - 1) = "" And wz(ch + 3, cz + 3) = "" Then
                If wz(ch - 2, cz - 2) = "" Then
                   jls(47) = jls(47) + 1
                End If
             End If
          End If
       End If
       If ch >= 3 And ch <= dhz - 3 Then
          If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
             If wz(ch + 1, cz - 1) = "" And wz(ch - 3, cz + 3) = "" Then
                If wz(ch + 2, cz - 2) = "" Then
                   jls(48) = jls(48) + 1
                End If
             End If
          End If
       End If
    End If
    If ch = 1 Then
       If wz(ch, cz) = wz(ch + i, cz) And wz(ch + i, cz) <> "" Then
          If wz(ch - 1, cz) = "" And wz(ch + 3, cz) = "" Then
             If wz(ch + 4, cz) = "" Then
                jls(17) = jls(17) + 1
             End If
          End If
       End If
       If cz >= 1 And cz <= dhz - 5 Then
          If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
             If wz(ch - 1, cz - 1) = "" And wz(ch + 3, cz + 3) = "" Then
                If wz(ch + 4, cz + 4) = "" Then
                   jls(18) = jls(18) + 1
                End If
             End If
          End If
       End If
       If cz >= 4 And cz <= dhz - 2 Then
          If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
             If wz(ch - 1, cz + 1) = "" And wz(ch + 3, cz - 3) = "" Then
                If wz(ch + 4, cz - 4) = "" Then
                   jls(19) = jls(19) + 1
                End If
             End If
          End If
       End If
    End If
    If ch = 2 Then
       If wz(ch - 1, cz) = wz(ch - 1 + i, cz) And wz(ch - 1 + i, cz) <> "" Then
          If wz(ch - 2, cz) = "" And wz(ch + 2, cz) = "" Then
             If wz(ch + 3, cz) = "" Then
                jls(20) = jls(20) + 1
             End If
          End If
       End If
       If cz >= 2 And cz <= dhz - 4 Then
          If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
             If wz(ch - 2, cz - 2) = "" And wz(ch + 2, cz + 2) = "" Then
                If wz(ch + 3, cz + 3) = "" Then
                   jls(21) = jls(21) + 1
                End If
             End If
          End If
       End If
       If cz >= 3 And cz <= dhz - 5 Then
          If wz(ch - 1, cz + 1) = wz(ch - 1 + i, cz + 1 - i) And wz(ch - 1 + i, cz + 1 - i) <> "" Then
             If wz(ch - 2, cz + 2) = "" And wz(ch + 2, cz - 2) = "" Then
                If wz(ch + 3, cz - 3) = "" Then
                   jls(22) = jls(22) + 1
                End If
             End If
          End If
       End If
    End If
    If cz = 1 Then
       If wz(ch, cz) = wz(ch, cz + i) And wz(ch, cz + i) <> "" Then
          If wz(ch, cz - 1) = "" And wz(ch, cz + 3) = "" Then
             If wz(ch, cz + 4) = "" Then
                jls(23) = jls(23) + 1
             End If
          End If
       End If
       If ch > 1 And ch <= dhz - 5 Then
          If wz(ch, cz) = wz(ch + i, cz + i) And wz(ch + i, cz + i) <> "" Then
             If wz(ch - 1, cz - 1) = "" And wz(ch + 3, cz + 3) = "" Then
                If wz(ch + 4, cz + 4) = "" Then
                   jls(24) = jls(24) + 1
                End If
             End If
          End If
       End If
       If ch >= 4 And ch <= dhz - 2 Then
          If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
             If wz(ch + 1, cz - 1) = "" And wz(ch - 3, cz + 3) = "" Then
                If wz(ch - 4, cz + 4) = "" Then
                   jls(25) = jls(25) + 1
                End If
             End If
          End If
       End If
    End If
    If cz = 2 Then
       If wz(ch, cz - 1) = wz(ch, cz - 1 + i) And wz(ch, cz - 1 + i) <> "" Then
          If wz(ch, cz - 2) = "" And wz(ch, cz + 2) = "" Then
             If wz(ch, cz + 3) = "" Then
                jls(26) = jls(26) + 1
             End If
          End If
       End If
       If ch > 2 And ch <= dhz - 4 Then
          If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
             If wz(ch - 2, cz - 2) = "" And wz(ch + 2, cz + 2) = "" Then
                If wz(ch + 3, cz + 3) = "" Then
                   jls(27) = jls(27) + 1
                End If
             End If
          End If
       End If
       If ch >= 3 And ch <= dhz - 3 Then
          If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
             If wz(ch + 2, cz - 2) = "" And wz(ch - 2, cz + 2) = "" Then
                If wz(ch - 3, cz + 3) = "" Then
                   jls(28) = jls(28) + 1
                End If
             End If
          End If
       End If
    End If
    If ch = dhz - 2 Then
       If wz(ch, cz) = wz(ch - i, cz) And wz(ch - i, cz) <> "" Then
          If wz(ch + 1, cz) = "" And wz(ch - 3, cz) = "" Then
             If wz(ch - 4, cz) = "" Then
                jls(29) = jls(29) + 1
             End If
          End If
       End If
       If cz >= 1 And cz <= dhz - 5 Then
          If wz(ch, cz) = wz(ch - i, cz + i) And wz(ch - i, cz + i) <> "" Then
             If wz(ch + 1, cz - 1) = "" And wz(ch - 3, cz + 3) = "" Then
                If wz(ch - 4, cz + 4) = "" Then
                   jls(30) = jls(30) + 1
                End If
             End If
          End If
       End If
       If cz >= 4 And cz <= dhz - 2 Then
          If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
             If wz(ch + 1, cz + 1) = "" And wz(ch - 3, cz - 3) = "" Then
                If wz(ch - 4, cz - 4) = "" Then
                   jls(31) = jls(31) + 1
                End If
             End If
          End If
       End If
    End If
    If ch = dhz - 3 Then
       If wz(ch - 1, cz) = wz(ch - 1 + i, cz) And wz(ch - 1 + i, cz) <> "" Then
          If wz(ch - 2, cz) = "" And wz(ch + 2, cz) = "" Then
             If wz(ch - 3, cz) = "" Then
                jls(32) = jls(32) + 1
             End If
          End If
       End If
       If cz >= 2 And cz <= dhz - 4 Then
          If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
             If wz(ch + 2, cz - 2) = "" And wz(ch - 2, cz + 2) = "" Then
                If wz(ch - 3, cz + 3) = "" Then
                   jls(33) = jls(33) + 1
                End If
             End If
          End If
       End If
       If cz >= 3 And cz <= dhz - 3 Then
          If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
             If wz(ch - 2, cz - 2) = "" And wz(ch + 2, cz + 2) = "" Then
                If wz(ch - 3, cz - 3) = "" Then
                   jls(34) = jls(34) + 1
                End If
             End If
          End If
       End If
    End If
    If cz = dhz - 2 Then
       If wz(ch, cz) = wz(ch, cz - i) And wz(ch, cz - i) <> "" Then
          If wz(ch, cz + 1) = "" And wz(ch, cz - 3) = "" Then
             If wz(ch, cz - 4) = "" Then
                jls(35) = jls(35) + 1
             End If
          End If
       End If
       If ch >= 4 And ch <= dhz - 2 Then
          If wz(ch, cz) = wz(ch - i, cz - i) And wz(ch - i, cz - i) <> "" Then
             If wz(ch + 1, cz + 1) = "" And wz(ch - 3, cz - 3) = "" Then
                If wz(ch - 4, cz - 4) = "" Then
                   jls(36) = jls(36) + 1
                End If
             End If
          End If
       End If
       If ch >= 1 And ch <= dhz - 5 Then
          If wz(ch, cz) = wz(ch + i, cz - i) And wz(ch + i, cz - i) <> "" Then
             If wz(ch - 1, cz + 1) = "" And wz(ch + 3, cz - 3) = "" Then
                If wz(ch + 4, cz - 4) = "" Then
                   jls(37) = jls(37) + 1
                End If
             End If
          End If
       End If
    End If
    If cz = dhz - 3 Then
       If wz(ch, cz - 1) = wz(ch, cz - 1 + i) And wz(ch, cz - 1 + i) <> "" Then
          If wz(ch, cz - 2) = "" And wz(ch, cz + 2) = "" Then
             If wz(ch, cz - 3) = "" Then
                jls(38) = jls(38) + 1
             End If
          End If
       End If
       If ch >= 2 And ch <= dhz - 4 Then
          If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
             If wz(ch + 2, cz - 2) = "" And wz(ch - 2, cz + 2) = "" Then
                If wz(ch + 3, cz - 3) = "" Then
                   jls(39) = jls(39) + 1
                End If
             End If
          End If
       End If
       If ch >= 3 And ch <= dhz - 3 Then
          If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
             If wz(ch - 2, cz - 2) = "" And wz(ch + 2, cz + 2) = "" Then
                If wz(ch - 3, cz - 3) = "" Then
                   jls(40) = jls(40) + 1
                End If
             End If
          End If
        End If
      End If
    Next i
    For j = 1 To 48
        If jls(j) = 2 Then
           three = 1
           Exit For
        End If
    Next j
    End If
    If three = True Then Exit For
  Next cz
  If three = True Then Exit For
Next ch
If three = True Then
   Dim jz1!, jz2!
   Select Case 2
        Case jls(1)
             If bo = True Then
                If wz(ch - 2, cz) = "" Then
                   h = ch - 1: z = cz
                ElseIf wz(ch + 4, cz) = "" Then
                       h = ch + 3: z = cz
                End If
             ElseIf bo = False Then
                    wz(ch - 1, cz) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch - 1, cz) = ""
                    wz(ch + 3, cz) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch + 3, cz) = ""
                    If jz1 >= jz2 Then
                       h = ch - 1: z = cz
                    Else
                       h = ch + 3: z = cz
                    End If
             End If
        Case jls(2)
             If bo = True Then
                If wz(ch, cz - 2) = "" Then
                   h = ch: z = cz - 1
                ElseIf wz(ch, cz + 4) = "" Then
                       h = ch: z = cz + 3
                End If
             ElseIf bo = False Then
                     wz(ch, cz - 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch, cz - 1) = ""
                    wz(ch, cz + 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch, cz + 3) = ""
                    If jz1 >= jz2 Then
                       h = ch: z = cz - 1
                    Else
                       h = ch: z = cz + 3
                    End If
             End If
        Case jls(3)
             If bo = True Then
                If wz(ch + 2, cz) = "" Then
                   h = ch + 1: z = cz
                ElseIf wz(ch - 4, cz) = "" Then
                       h = ch - 3: z = cz
                End If
             Else
                    wz(ch + 1, cz) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch + 1, cz) = ""
                    wz(ch - 3, cz) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch - 3, cz) = ""
                    If jz1 >= jz2 Then
                       h = ch + 1: z = cz
                    Else
                       h = ch - 3: z = cz
                    End If
             End If
        Case jls(4)
             If bo = True Then
                If wz(ch, cz + 2) = "" Then
                   h = ch: z = cz + 1
                ElseIf wz(ch, cz - 4) = "" Then
                       h = ch: z = cz - 3
                End If
             Else
                    wz(ch, cz + 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch, cz + 1) = ""
                    wz(ch, cz - 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch, cz - 3) = ""
                    If jz1 >= jz2 Then
                       h = ch: z = cz + 1
                    Else
                       h = ch: z = cz - 3
                    End If
             End If
        Case jls(5)
             If bo = True Then
                If wz(ch - 2, cz - 2) = "" Then
                   h = ch - 1: z = cz - 1
                ElseIf wz(ch + 4, cz + 4) = "" Then
                       h = ch + 3: z = cz + 3
                End If
             Else
                    wz(ch - 1, cz - 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch - 1, cz - 1) = ""
                    wz(ch + 3, cz + 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch + 3, cz + 3) = ""
                    If jz1 >= jz2 Then
                       h = ch - 1: z = cz - 1
                    Else
                       h = ch + 3: z = cz + 3
                    End If
             End If
        Case jls(6)
             If bo = True Then
                If wz(ch + 2, cz - 2) = "" Then
                   h = ch + 1: z = cz - 1
                ElseIf wz(ch - 4, cz + 4) = "" Then
                       h = ch - 3: z = cz + 3
                End If
             Else
                    wz(ch + 1, cz - 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch + 1, cz - 1) = ""
                    wz(ch - 3, cz + 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch - 3, cz + 3) = ""
                    If jz1 >= jz2 Then
                       h = ch + 1: z = cz - 1
                    Else
                       h = ch - 3: z = cz + 3
                    End If
             End If
        Case jls(7)
             If bo = True Then
                If wz(ch + 2, cz + 2) = "" Then
                   h = ch + 1: z = cz + 1
                ElseIf wz(ch - 4, cz - 4) = "" Then
                       h = ch - 3: z = cz - 3
                End If
             Else
                    wz(ch + 1, cz + 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch + 1, cz + 1) = ""
                    wz(ch - 3, cz - 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch - 3, cz - 3) = ""
                    If jz1 >= jz2 Then
                       h = ch + 1: z = cz + 1
                    Else
                       h = ch - 3: z = cz - 3
                    End If
             End If
        Case jls(8)
             If bo = True Then
                If wz(ch - 2, cz + 2) = "" Then
                   h = ch - 1: z = cz + 1
                ElseIf wz(ch + 4, cz - 4) = "" Then
                       h = ch + 3: z = cz - 3
                End If
             Else
                    wz(ch - 1, cz + 1) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch - 1, cz + 1) = ""
                    wz(ch + 3, cz - 3) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch + 3, cz - 3) = ""
                    If jz1 >= jz2 Then
                       h = ch - 1: z = cz + 1
                    Else
                       h = ch + 3: z = cz - 3
                    End If
             End If
        Case jls(9)
             If bo = True Then
                If wz(ch - 3, cz - 3) = "" Then
                   h = ch - 2: z = cz - 2
                ElseIf wz(ch + 3, cz + 3) = "" Then
                       h = ch + 2: z = cz + 2
                End If
             Else
                    wz(ch - 2, cz - 2) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch - 2, cz - 2) = ""
                    wz(ch + 2, cz + 2) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch + 2, cz + 2) = ""
                    If jz1 >= jz2 Then
                       h = ch - 2: z = cz - 2
                    Else
                       h = ch + 2: z = cz + 2
                    End If
             End If
        Case jls(10)
             If bo = True Then
                If wz(ch + 3, cz - 3) = "" Then
                   h = ch + 2: z = cz - 2
                ElseIf wz(ch - 3, cz + 3) = "" Then
                       h = ch - 2: z = cz + 2
                End If
             Else
                    wz(ch + 2, cz - 2) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch + 2, cz - 2) = ""
                    wz(ch - 2, cz + 2) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch - 2, cz + 2) = ""
                    If jz1 >= jz2 Then
                       h = ch + 2: z = cz - 2
                    Else
                       h = ch - 2: z = cz + 2
                    End If
             End If
        Case jls(11)
             If bo = True Then
                If wz(ch - 3, cz) = "" Then
                   h = ch - 2: z = cz
                ElseIf wz(ch + 3, cz) = "" Then
                       h = ch + 2: z = cz
                End If
             Else
                    wz(ch - 2, cz) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch - 2, cz) = ""
                    wz(ch + 2, cz) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch + 2, cz) = ""
                    If jz1 >= jz2 Then
                       h = ch - 2: z = cz
                    Else
                       h = ch + 2: z = cz
                    End If
             End If
        Case jls(12)
             If bo = True Then
                If wz(ch, cz - 3) = "" Then
                   h = ch: z = cz - 2
                ElseIf wz(ch, cz + 3) = "" Then
                       h = ch: z = cz + 2
                End If
             Else
                    wz(ch, cz - 2) = "����"
                    jz1 = estimate(wz(), Not bo)
                    wz(ch, cz - 2) = ""
                    wz(ch, cz + 2) = "����"
                    jz2 = estimate(wz(), Not bo)
                    wz(ch, cz + 2) = ""
                    If jz1 >= jz2 Then
                       h = ch: z = cz - 2
                    Else
                       h = ch: z = cz + 2
                    End If
             End If
        Case jls(13)
             h = ch - 1: z = cz
        Case jls(14)
             h = ch + 1: z = cz
        Case jls(15)
             h = ch: z = cz + 1
        Case jls(16)
             h = ch: z = cz - 1
        Case jls(17)
             h = ch + 3: z = cz
        Case jls(18)
             h = ch + 3: z = cz + 3
        Case jls(19)
             h = ch + 3: z = cz - 3
        Case jls(20)
             h = ch + 2: z = cz
        Case jls(21)
             h = ch + 2: z = cz + 2
        Case jls(22)
             h = ch + 2: z = cz - 2
        Case jls(23)
             h = ch: z = cz + 3
        Case jls(24)
             h = ch + 3: z = cz + 3
        Case jls(25)
             h = ch - 3: z = cz + 3
        Case jls(26)
             h = ch: z = cz + 2
        Case jls(27)
             h = ch + 2: z = cz + 2
        Case jls(28)
             h = ch - 2: z = cz + 2
        Case jls(29)
             h = ch - 3: z = cz
        Case jls(30)
             h = ch - 3: z = cz + 3
        Case jls(31)
             h = ch - 3: z = cz - 3
        Case jls(32)
             h = ch - 2: z = cz
        Case jls(33)
             h = ch - 2: z = cz + 2
        Case jls(34)
             h = ch - 2: z = cz - 2
        Case jls(35)
             h = ch: z = cz - 3
        Case jls(36)
             h = ch - 3: z = cz - 3
        Case jls(37)
             h = ch + 3: z = cz - 3
        Case jls(38)
             h = ch: z = cz - 2
        Case jls(39)
             h = ch + 2: z = cz - 2
        Case jls(40)
             h = ch - 2: z = cz - 2
        Case jls(41)
             h = ch - 1: z = cz - 1
        Case jls(42)
             h = ch - 1: z = cz + 1
        Case jls(43)
             h = ch + 1: z = cz - 1
        Case jls(44)
             h = ch + 1: z = cz + 1
        Case jls(45)
             h = ch + 1: z = cz + 1
        Case jls(46)
             h = ch - 1: z = cz + 1
        Case jls(47)
             h = ch - 1: z = cz - 1
        Case jls(48)
             h = ch + 1: z = cz - 1
    End Select
End If
End Function

Public Function four%(wz$(), h%, z%, ByVal qzs As Boolean) '�ж��Լ������������������
Dim jls%(1 To 26)       '����ú���Ϊ�棬����������������ͷ������һ����λ��
Dim ch%, cz%, rt$       'Ҳ������Ӯ�����͡���ʱ���h��zΪ�������꣬�˵�������Ӯ��
four = 0            'qzsΪ���������������������Ϊ��������������
If qzs = True Then
   rt = "����"
Else
   rt = "����"
End If
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        For m = 1 To 26
            jls(m) = 0
        Next m
            If i >= 1 And i <= dhz - 5 And wz(i, j) = rt Then   '1-4���ң��ϣ����ϣ����ϱȽ���ǰ�����޿�λ
               If wz(i, j) = wz(i + 1, j) And wz(i, j) = wz(i + 2, j) And wz(i, j) = wz(i + 3, j) And wz(i, j) <> "" Then
                  If wz(i - 1, j) = "" Or wz(i + 4, j) = "" Then
                     jls(1) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If j >= 1 And j <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i, j + 1) And wz(i, j) = wz(i, j + 2) And wz(i, j) = wz(i, j + 3) And wz(i, j) <> "" Then
                  If wz(i, j - 1) = "" Or wz(i, j + 4) = "" Then
                     jls(2) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If i >= 1 And j >= 1 And i <= dhz - 5 And j <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = wz(i + 2, j + 2) And wz(i, j) = wz(i + 3, j + 3) And wz(i, j) <> "" Then
                  If wz(i - 1, j - 1) = "" Or wz(i + 4, j + 4) = "" Then
                     jls(3) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If i >= 4 And j >= 1 And i <= dhz - 2 And j <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i - 1, j + 1) And wz(i, j) = wz(i - 2, j + 2) And wz(i, j) = wz(i - 3, j + 3) And wz(i, j) <> "" Then
                  If wz(i + 1, j - 1) = "" Or wz(i - 4, j + 4) = "" Then
                     jls(4) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If i = dhz - 4 And wz(i, j) = rt Then         '���Һ���ȽϿ�������޿�λ
               If wz(i, j) = wz(i + 1, j) And wz(i, j) = wz(i + 2, j) And wz(i, j) = wz(i + 3, j) And wz(i, j) <> "" Then
                  If wz(i - 1, j) = "" Then
                     jls(5) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If i = 0 And wz(i, j) = rt Then            '���Һ���ȽϿ��ұ����޿�λ
               If wz(i, j) = wz(i + 1, j) And wz(i, j) = wz(i + 2, j) And wz(i, j) = wz(i + 3, j) And wz(i, j) <> "" Then
                  If wz(i + 4, j) = "" Then
                     jls(6) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If j = dhz - 4 And wz(i, j) = rt Then         '��������ȽϿ��·����޿�λ
               If wz(i, j) = wz(i, j + 1) And wz(i, j) = wz(i, j + 2) And wz(i, j) = wz(i, j + 3) And wz(i, j) <> "" Then
                  If wz(i, j - 1) = "" Then
                     jls(7) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If j = 0 And wz(i, j) = rt Then            '��������ȽϿ��Ϸ����޿�λ
               If wz(i, j) = wz(i, j + 1) And wz(i, j) = wz(i, j + 2) And wz(i, j) = wz(i, j + 3) And wz(i, j) <> "" Then
                  If wz(i, j + 4) = "" Then
                     jls(8) = 1
                     ch = i: cz = j
                  End If
               End If
            End If
            If (i = dhz - 4 Or j = dhz - 4) And wz(i, j) = rt Then '����б�ȽϿ��������޿�λ������ݣ�
               If i <= 16 And j = 16 And i >= 1 Then
                  If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = wz(i + 2, j + 2) And wz(i, j) = wz(i + 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i - 1, j - 1) = "" Then
                        jls(9) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
               If i = dhz - 4 And j <= dhz - 4 And j >= 1 And wz(i, j) = rt Then
                  If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = wz(i + 2, j + 2) And wz(i, j) = wz(i + 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i - 1, j - 1) = "" Then
                        jls(10) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If i = 0 Or j = 0 Then  '���ϱȽϿ��������޿�λ
               If i <= dhz - 5 And j <= dhz - 5 And wz(i, j) = rt Then
                  If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = wz(i + 2, j + 2) And wz(i, j) = wz(i + 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i + 4, j + 4) = "" Then
                        jls(11) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If j = dhz - 4 And wz(i, j) = rt Then         '����б�ȽϿ��������޿�λ
               If i <= dhz - 2 And i >= 3 Then
                  If wz(i, j) = wz(i - 1, j + 1) And wz(i, j) = wz(i - 2, j + 2) And wz(i, j) = wz(i - 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i + 1, j - 1) = "" Then
                        jls(12) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If i = 3 And wz(i, j) = rt Then             '����б�ȽϿ��������޿�λ
               If j <= dhz - 4 And j >= 1 Then
                  If wz(i, j) = wz(i - 1, j + 1) And wz(i, j) = wz(i - 2, j + 2) And wz(i, j) = wz(i - 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i + 1, j - 1) = "" Then
                        jls(13) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If i = dhz - 1 Or j = 0 Then '����б�ȽϿ��������޿�λ
               If i >= 4 And j <= dhz - 5 And wz(i, j) = rt Then
                  If wz(i, j) = wz(i - 1, j + 1) And wz(i, j) = wz(i - 2, j + 2) And wz(i, j) = wz(i - 3, j + 3) And wz(i, j) <> "" Then
                     If wz(i - 4, j + 4) = "" Then
                        jls(14) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If i <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i + 4, j) And wz(i, j) <> "" Then
                  If wz(i, j) = wz(i + 1, j) And wz(i, j) = wz(i + 2, j) Then
                     If wz(i + 3, j) = "" Then
                        jls(15) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j) And wz(i + 2, j) = wz(i, j) Then
                     If wz(i + 1, j) = "" Then
                        jls(16) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j) And wz(i, j) = wz(i + 1, j) Then
                     If wz(i + 2, j) = "" Then
                        jls(17) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If j <= dhz - 5 And wz(i, j) = rt Then
                If wz(i, j) = wz(i, j + 4) And wz(i, j) <> "" Then
                  If wz(i, j) = wz(i, j + 1) And wz(i, j) = wz(i, j + 2) Then
                     If wz(i, j + 3) = "" Then
                        jls(18) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i, j + 3) And wz(i, j + 2) = wz(i, j) Then
                     If wz(i, j + 1) = "" Then
                        jls(19) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i, j + 3) And wz(i, j) = wz(i, j + 1) Then
                     If wz(i, j + 2) = "" Then
                        jls(20) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If i <= dhz - 5 And j <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i + 4, j + 4) And wz(i, j) <> "" Then
                  If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = wz(i + 2, j + 2) Then
                     If wz(i + 3, j + 3) = "" Then
                        jls(21) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j + 3) And wz(i + 2, j + 2) = wz(i, j) Then
                     If wz(i + 1, j + 1) = "" Then
                        jls(22) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j + 3) And wz(i, j) = wz(i + 1, j + 1) Then
                     If wz(i + 2, j + 2) = "" Then
                        jls(23) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
            If j >= 4 And i <= dhz - 5 And wz(i, j) = rt Then
               If wz(i, j) = wz(i + 4, j - 4) And wz(i, j) <> "" Then
                  If wz(i, j) = wz(i + 1, j - 1) And wz(i, j) = wz(i + 2, j - 2) Then
                     If wz(i + 3, j - 3) = "" Then
                        jls(24) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j - 3) And wz(i + 2, j - 2) = wz(i, j) Then
                     If wz(i + 1, j - 1) = "" Then
                        jls(25) = 1
                        ch = i: cz = j
                     End If
                  End If
                  If wz(i, j) = wz(i + 3, j - 3) And wz(i, j) = wz(i + 1, j - 1) Then
                     If wz(i + 2, j - 2) = "" Then
                        jls(26) = 1
                        ch = i: cz = j
                     End If
                  End If
               End If
            End If
        For l = 1 To 26
            If jls(l) = 1 Then
               four = 1
               Exit For
            End If
        Next l
        If four = 1 Then Exit For
    Next j
    If four = 1 Then Exit For
Next i
If four = 1 Then
   Select Case 1
       Case jls(1)
            If wz(ch - 1, cz) = "" Then
               h = ch - 1
               z = cz
            ElseIf wz(ch + 4, cz) = "" Then
                   h = ch + 4
                   z = cz
            End If
            If wz(ch - 1, cz) = "" And wz(ch + 4, cz) = "" Then
               four = 2
            End If
       Case jls(2)
            If wz(ch, cz - 1) = "" Then
               h = ch
               z = cz - 1
            ElseIf wz(ch, cz + 4) = "" Then
                   h = ch
                   z = cz + 4
            End If
            If wz(ch, cz - 1) = "" And wz(ch, cz + 4) = "" Then
               four = 2
            End If
       Case jls(3)
            If wz(ch - 1, cz - 1) = "" Then
               h = ch - 1
               z = cz - 1
            ElseIf wz(ch + 4, cz + 4) = "" Then
                   h = ch + 4
                   z = cz + 4
            End If
            If wz(ch - 1, cz - 1) = "" And wz(ch + 4, cz + 4) = "" Then
               four = 2
            End If
       Case jls(4)
            If wz(ch + 1, cz - 1) = "" Then
               h = ch + 1
               z = cz - 1
            ElseIf wz(ch - 4, cz + 4) = "" Then
                   h = ch - 4
                   z = cz + 4
            End If
            If wz(ch + 1, cz - 1) = "" And wz(ch - 4, cz + 4) = "" Then
               four = 2
            End If
       Case jls(5)
            h = ch - 1
            z = cz
       Case jls(6)
            h = ch + 4
            z = cz
       Case jls(7)
            h = ch
            z = cz - 1
       Case jls(8)
            h = ch
            z = cz + 4
       Case jls(9)
            h = ch - 1
            z = cz - 1
       Case jls(10)
            h = ch - 1
            z = cz - 1
       Case jls(11)
            h = ch + 4
            z = cz + 4
       Case jls(12)
            h = ch + 1
            z = cz - 1
       Case jls(13)
            h = ch - 4
            z = cz + 4
       Case jls(14)
            h = ch - 4
            z = cz + 4
       Case jls(15)
            h = ch + 3
            z = cz
       Case jls(16)
            h = ch + 1
            z = cz
       Case jls(17)
            h = ch + 2
            z = cz
       Case jls(18)
            h = ch
            z = cz + 3
       Case jls(19)
            h = ch
            z = cz + 1
       Case jls(20)
            h = ch
            z = cz + 2
       Case jls(21)
            h = ch + 3
            z = cz + 3
       Case jls(22)
            h = ch + 1
            z = cz + 1
       Case jls(23)
            h = ch + 2
            z = cz + 2
       Case jls(24)
            h = ch + 3
            z = cz - 3
       Case jls(25)
            h = ch + 1
            z = cz - 1
       Case jls(26)
            h = ch + 2
            z = cz - 2
   End Select
End If
End Function

Public Sub sl(ByVal fd$)
Dim cu As save, jlh%
az = MsgBox("�Ƿ񱣴������", 36, "��ʾ")
If az = vbYes Then
   com.CancelError = True
   On Error GoTo errhandler
   com.ShowSave
   If com.FileName <> "" Then
       Open com.FileName & "-" & fd & ".lsl" For Random As #1 Len = Len(cu)
           jlh = 1
           Do Until hz(jlh) = "" And bz(jlh) = ""
              cu.zbh = hz(jlh): cu.ysh = ys1
              cu.zbb = bz(jlh): cu.ysb = ys2
              cu.sjh = sjh: cu.sjb = sjb
              If hz(jlh) = "" Then
                    cu.sjb = Val(bstr(md, dhz))
              End If
              If bz(jlh) = "" Then
                 cu.sjh = Val(bstr(md, dhz))
              End If
              Put #1, jlh, cu
              jlh = jlh + 1
           Loop
           If hz(jlh - 1) <> "" And bz(jlh - 1) <> "" Then
              If slz = True Then
                 cu.zbh = "": cu.ysh = 0
                 cu.zbb = "1234": cu.ysb = 0
                 cu.sjh = 0: cu.sjb = Val(bstr(md, dhz))
                 Put #1, jlh, cu
              ElseIf slz = False Then
                     cu.zbh = "1234": cu.ysh = 0
                     cu.zbb = "": cu.ysb = 0
                     cu.sjh = Val(bstr(md, dhz)): cu.sjb = 0
                     Put #1, jlh, cu
              End If
           End If
       Close #1
       MsgBox "�����ѱ�����" & com.FileName & ".lsl", 48, "��ʾ"
   End If
errhandler:
End If
tsbc = True
Picture2.Enabled = False
Timerh.Enabled = False
Timerb.Enabled = False
If fd <> "" Then
   On Error Resume Next
   Kill App.Path & "���������.lsl"
       Open App.Path & "���������.lsl" For Random As #1 Len = Len(cu)
           FileName = "": jlh = 1
           Do Until hz(jlh) = "" And bz(jlh) = ""
              cu.zbh = hz(jlh): cu.ysh = ys1
              cu.zbb = bz(jlh): cu.ysb = ys2
              cu.sjh = sjh: cu.sjb = sjb
              If hz(jlh) = "" Then
                    cu.sjb = Val(bstr(md, dhz))
              End If
              If bz(jlh) = "" Then
                 cu.sjh = Val(bstr(md, dhz))
              End If
              Put #1, jlh, cu
              jlh = jlh + 1
           Loop
           If hz(jlh - 1) <> "" And bz(jlh - 1) <> "" Then
              If slz = True Then
                 cu.zbh = "": cu.ysh = 0
                 cu.zbb = "1234": cu.ysb = 0
                 cu.sjh = 0: cu.sjb = Val(bstr(md, dhz))
                 Put #1, jlh, cu
              ElseIf slz = False Then
                     cu.zbh = "1234": cu.ysh = 0
                     cu.zbb = "": cu.ysb = 0
                     cu.sjh = Val(bstr(md, dhz)): cu.sjb = 0
                     Put #1, jlh, cu
              End If
           End If
       Close #1
End If
Dim yl As dlm, i%
Open App.Path & "zcb.lsn" For Random As #1 Len = Len(yl)
     For i = 1 To LOF(1) / Len(yl)
         Get #1, i, yl
         If yl.mz = dl.mz Then
            Exit For
         End If
     Next i
     If md = 1 Then
        If slz = True Then
           yl.drb.bs_w = bsb + yl.drb.bs_w
           yl.drb.sj_w = sjb + yl.drb.sj_w
           yl.drh.bs_f = bsh + yl.drh.bs_f
           yl.drh.sj_f = sjh + yl.drh.sj_f
           yl.drb.win_ = yl.drb.win_ + 1
           yl.drh.fail = yl.drh.fail + 1
        Else
            yl.drh.bs_w = bsh + yl.drh.bs_w
            yl.drh.sj_w = sjh + yl.drh.sj_w
            yl.drb.bs_f = bsb + yl.drb.bs_f
            yl.drb.sj_f = sjb + yl.drb.sj_f
            yl.drh.win_ = yl.drh.win_ + 1
            yl.drb.fail = yl.drb.fail + 1
        End If
     ElseIf md = 2 Then
            If slz = True Then
               yl.rj.bs_w = bsb + yl.rj.bs_w
               yl.rj.sj_w = sjb + yl.rj.sj_w
               yl.rj.win_ = yl.rj.win_ + 1
            Else
                yl.rj.bs_f = bsb + yl.rj.bs_f
                yl.rj.sj_f = sjb + yl.rj.sj_f
                yl.rj.fail = yl.rj.fail + 1
            End If
     ElseIf md = 3 Then
            If slz = True Then
               yl.wl.bs_f = bsh + yl.wl.bs_f
               yl.wl.sj_f = sjh + yl.wl.sj_f
               yl.wl.fail = yl.wl.fail + 1
            Else
                yl.wl.bs_w = bsh + yl.wl.bs_w
                yl.wl.sj_w = sjh + yl.wl.sj_w
                yl.wl.win_ = yl.wl.win_ + 1
            End If
     End If
     Put #1, i, yl
     dl = yl
Close #1
End Sub

Public Function twoone(wz$(), ByRef h%, ByRef z%, ByVal qz As Boolean) As Boolean
Dim jls%(1 To 8), xh%, xz%, ty$
twoone = False
If qz = True Then
   ty = "����"
Else
   ty = "����"
End If
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If i >= 1 And i <= dhz - 5 And wz(i, j) = ty Then
           If wz(i, j) = wz(i + 3, j) And wz(i - 1, j) = "" And wz(i + 4, j) = "" Then
              If wz(i, j) = wz(i + 1, j) And wz(i + 2, j) = "" Then
                 jls(1) = 1: xh = i + 2: xz = j
              End If
              If wz(i, j) = wz(i + 2, j) And wz(i + 1, j) = "" Then
                 jls(2) = 1: xh = i + 1: xz = j
              End If
           End If
        End If
        If j >= 1 And j <= dhz - 5 And wz(i, j) = ty Then
           If wz(i, j) = wz(i, j + 3) And wz(i, j - 1) = "" And wz(i, j + 4) = "" Then
              If wz(i, j) = wz(i, j + 1) And wz(i, j + 2) = "" Then
                 jls(3) = 1: xh = i: xz = j + 2
              End If
              If wz(i, j) = wz(i, j + 2) And wz(i, j + 1) = "" Then
                 jls(4) = 1: xh = i: xz = j + 1
              End If
           End If
        End If
        If i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 And wz(i, j) = ty Then
           If wz(i, j) = wz(i + 3, j + 3) And wz(i - 1, j - 1) = "" And wz(i + 4, j + 4) = "" Then
              If wz(i, j) = wz(i + 1, j + 1) And wz(i + 2, j + 2) = "" Then
                 jls(5) = 1: xh = i + 2: xz = j + 2
              End If
              If wz(i, j) = wz(i + 2, j + 2) And wz(i + 1, j + 1) = "" Then
                 jls(6) = 1: xh = i + 1: xz = j + 1
              End If
           End If
        End If
        If i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 And wz(i, j) = ty Then
           If wz(i, j) = wz(i + 3, j - 3) And wz(i - 1, j + 1) = "" And wz(i + 4, j - 4) = "" Then
              If wz(i, j) = wz(i + 1, j - 1) And wz(i + 2, j - 2) = "" Then
                 jls(7) = 1: xh = i + 2: xz = j - 2
              End If
              If wz(i, j) = wz(i + 2, j - 2) And wz(i + 1, j - 1) = "" Then
                 jls(8) = 1: xh = i + 1: xz = j - 1
              End If
           End If
        End If
        For m = 1 To 8
            If jls(m) = 1 Then
               twoone = True
               Exit For
            End If
        Next m
        If twoone = True Then Exit For
    Next j
    If twoone = True Then Exit For
Next i
If twoone = True Then
   Dim jz1!, jz2!, jz3!
   Select Case 1
         Case jls(1)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 3, xz) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 3, xz) = ""
                     wz(xh + 2, xz) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 2, xz) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 3: z = xz
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 2: z = xz
                     End If
              End If
         Case jls(2)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 2, xz) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 2, xz) = ""
                     wz(xh + 3, xz) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 3, xz) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 2: z = xz
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 3: z = xz
                     End If
              End If
         Case jls(3)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh, xz - 3) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh, xz - 3) = ""
                     wz(xh, xz + 2) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh, xz + 2) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh: z = xz - 3
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh: z = xz + 2
                     End If
              End If
         Case jls(4)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh, xz - 2) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh, xz - 2) = ""
                     wz(xh, xz + 3) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh, xz + 3) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh: z = xz - 2
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh: z = xz + 3
                     End If
              End If
         Case jls(5)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 3, xz - 3) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 3, xz - 3) = ""
                     wz(xh + 2, xz + 2) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 2, xz + 2) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 3: z = xz - 3
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 2: z = xz + 2
                     End If
              End If
         Case jls(6)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 2, xz - 2) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 2, xz - 2) = ""
                     wz(xh + 3, xz + 3) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 3, xz + 3) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 2: z = xz - 2
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 3: z = xz + 3
                     End If
              End If
         Case jls(7)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 3, xz + 3) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 3, xz + 3) = ""
                     wz(xh + 2, xz - 2) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 2, xz - 2) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 3: z = xz + 3
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 2: z = xz - 2
                     End If
              End If
         Case jls(8)
              If qz = True Then
                 h = xh: z = xz
              ElseIf qz = False Then
                     wz(xh, xz) = "����"
                     jz1 = estimate(wz(), Not qz)
                     wz(xh, xz) = ""
                     wz(xh - 2, xz + 2) = "����"
                     jz2 = estimate(wz(), Not qz)
                     wz(xh - 2, xz + 2) = ""
                     wz(xh + 3, xz - 3) = "����"
                     jz3 = estimate(wz(), Not qz)
                     wz(xh + 3, xz - 3) = ""
                     If jz1 >= jz2 And jz1 >= jz3 Then
                        h = xh: z = xz
                     ElseIf jz2 >= jz1 And jz2 >= jz3 Then
                            h = xh - 2: z = xz + 2
                     ElseIf jz3 >= jz1 And jz3 >= jz2 Then
                            h = xh + 3: z = xz - 3
                     End If
              End If
   End Select
End If
End Function

Public Function autolz(wz$()) As String   '�Զ�����
Dim lz!(24, 24), ah%, az%, st$, qzs As Boolean, dkh$(), dfh$(), dkb$(), dfb$(), dgh$(), dgb$(), _
th As Boolean, sh As Boolean, tb As Boolean, sb As Boolean, rh As Boolean, rb As Boolean
qzs = True: sh = False: th = False: sb = False: tb = False: rh = False: rb = False
'//////////////////////////////////////////////////////////////////////////////
If four(wz(), ah, az, qzs) <> 0 Then '�������������������Ӯ�ñ���
   If Option1.Value = True And cljs.Checked = True Then
      If jsix(ah, az, True) = True Then
         lz(ah, az) = 1
      End If
   Else
       autolz = bstr(ah, az)
       Exit Function
   End If
ElseIf four(wz(), ah, az, Not qzs) <> 0 Then  '����Է�������������������
       autolz = bstr(ah, az)
       Exit Function
End If
'////////////////////////////////////////////////////////�������
If san(wz(), qzs, dfh()) = True Then     'ǿ����
   sh = True
   ssz = True
   st = compare(dfh(), dfh(), qzs)
   ssz = False
   If st <> "" Then
      If Option1.Value = True And sijs.Checked = True Then
         For i = 1 To Len(st) / 4
             Call jstr(Mid(st, 1 + (i - 1) * 4, 4), ah, az)
             lz(ah, az) = 1
         Next i
      ElseIf Option1.Value = False And sijs.Checked = True Then
          If Len(st) > 4 Then
             autolz = Left(st, 4)
             Exit Function
          Else
              autolz = st
              Exit Function
          End If
      End If
   End If
End If
If rsan(wz(), qzs, dgh()) = True Then    '������
   rh = True
   ggz = True
   st = compare(dgh(), dgh(), qzs)
   ggz = False
   If st <> "" Then
      If Option1.Value = True And sijs.Checked = True Then
         For i = 1 To Len(st) / 4
             Call jstr(Mid(st, 1 + (i - 1) * 4, 4), ah, az)
             lz(ah, az) = 1
         Next i
      ElseIf Option1.Value = False And sijs.Checked = True Then
          If Len(st) > 4 Then
             autolz = Left(st, 4)
             Exit Function
          Else
              autolz = st
              Exit Function
          End If
      End If
   End If
End If
If two(wz(), qzs, dkh()) = True Then th = True
If th = True And sh = True Then                 'ǿ����
   st = compare(dkh(), dfh(), qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If rh = True And sh = True Then                 'ǿ������
   rrz = True
   st = compare(dgh(), dfh(), qzs)
   rrz = False
   If st <> "" Then
      If Option1.Value = True And sijs.Checked = True Then
         For i = 1 To Len(st) / 4
             Call jstr(Mid(st, 1 + (i - 1) * 4, 4), ah, az)
             lz(ah, az) = 1
         Next i
      ElseIf Option1.Value = False And sijs.Checked = True Then
          If Len(st) > 4 Then
             autolz = Left(st, 4)
             Exit Function
          Else
              autolz = st
              Exit Function
          End If
      End If
   End If
End If
'/////////////////////////////////////////////////////��������������
If three(wz(), ah, az, qzs) = True Then
   autolz = bstr(ah, az)
   Exit Function
ElseIf twoone(wz(), ah, az, qzs) = True Then
   autolz = bstr(ah, az)
   Exit Function
ElseIf three(wz(), ah, az, Not qzs) = True Then
   autolz = bstr(ah, az)
   Exit Function
ElseIf twoone(wz(), ah, az, Not qzs) = True Then
       autolz = bstr(ah, az)
       Exit Function
End If
'//////////////////////////////////////////////////////������
If san(wz(), Not qzs, dfb()) = True Then        'ǿ����
   sb = True
   st = compare(dfb(), dfb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If rsan(wz(), Not qzs, dgb()) = True Then      '������
   rb = True
   st = compare(dgb(), dgb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If rb = True And sb = True Then             'ǿ������
   st = compare(dgb(), dfb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If two(wz(), Not qzs, dkb()) = True Then tb = True
If tb = True And sb = True Then                 'ǿ����
   st = compare(dkb(), dfb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
'////////////////////////////////////////////�������
If rh = True And th = True Then
   st = compare(dkh(), dgh(), qzs)
   If st <> "" Then
     If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If th = True Then
   st = compare(dkh(), dkh(), qzs)
   If st <> "" Then
      If Option1.Value = True And ssjs.Checked = True Then
         For i = 1 To Len(st) / 4
             Call jstr(Mid(st, 1 + (i - 1) * 4, 4), ah, az)
             lz(ah, az) = 1
         Next i
      ElseIf Option1.Value = False And ssjs.Checked = True Then
          If Len(st) > 4 Then
             autolz = Left(st, 4)
             Exit Function
          Else
              autolz = st
              Exit Function
          End If
      End If
   End If
End If
oot = False: thtw = False: tto = False: gyg = False
'/////////////////////////////////////////������
If rb = True And tb = True Then
   st = compare(dkb(), dgb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
If tb = True Then
   st = compare(dkb(), dkb(), Not qzs)
   If st <> "" Then
      If Len(st) > 4 Then
         autolz = Left(st, 4)
         Exit Function
      Else
          autolz = st
          Exit Function
      End If
   End If
End If
'/////////////////////////////////////////////////�������ӵ�
Dim big!, tr$, shen$
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        wez(i, j) = wz(i, j)
    Next j
Next i
If jgx.Enabled = False Then
   shen = "����"
   qzs = True
ElseIf fsx.Enabled = False Then
       shen = "����"
       qzs = False
End If
big = -100000
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If wez(i, j) <> "" Then
           For a = i - 2 To i + 2
               For b = j - 2 To j + 2
                   If a >= 0 And a <= dhz - 1 And b >= 0 And b <= dhz - 1 Then
                      If wez(a, b) = "" And lz(a, b) = 0 Then
                         wez(a, b) = shen
                         lz(a, b) = estimate(wez(), qzs)
                         wez(a, b) = ""
                         If lz(a, b) > big Then
                            big = lz(a, b)
                            tr = bstr(a, b)
                         End If
                      End If
                   End If
               Next b
           Next a
        End If
    Next j
Next i
autolz = tr
End Function
Public Function bstr(ByVal HS%, ByVal zs%) As String
Dim m$, n$                     '��������Ϣ���ܳ�һ���ַ��������ڴ��䡣
If HS <= 9 And HS >= 0 Then
   m = "9" & HS
ElseIf HS >= 10 And HS <= 25 Then
       m = HS
End If
If zs <= 9 And zs >= 0 Then
   n = "9" & zs
ElseIf zs >= 10 And zs <= 25 Then
       n = zs
End If
bstr = Trim(Str(m & n))
End Function

Public Sub jstr(ByVal tr$, ByRef hf%, ByRef zs%)
If Left(tr, 1) = "9" Then   '�����ַ����е�������Ϣ
   hf = Val(Mid(tr, 2, 1))
ElseIf Left(tr, 1) = "1" Or Left(tr, 1) = "2" Then
       hf = Val(Mid(tr, 1, 2))
End If
If Mid(tr, 3, 1) = "9" Then
   zs = Val(Mid(tr, 4, 1))
ElseIf Mid(tr, 3, 1) = "1" Or Mid(tr, 3, 1) = "2" Then
       zs = Val(Mid(tr, 3, 2))
End If
End Sub

Public Function estimate!(ew$(), ByVal qw As Boolean)      '��������
Dim vh!(24), vz!(24), lu!(1 To 84), m$, n$, su!
Static cs%
If qw = True Then
   m = "����"
   n = "����"
Else
   m = "����"
   n = "����"
End If
For i = 0 To dhz - 1
    For q = 0 To 24
        vz(q) = 0
    Next q
    For j = 0 To dhz - 1
        For k = 1 To 84
            lu(k) = 0
        Next k
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then '����һ��
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = "" And ew(i + 3, j) = "" Then
                 lu(1) = 10
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And lu(1) = 0 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = "" And ew(i + 2, j) = "" Then
                     lu(2) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And lu(1) = 0 And lu(2) = 0 Then
               If ew(i - 3, j) = "" And ew(i + 2, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i - 2, j) = "" And ew(i + 1, j) = "" Then
                     lu(3) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 4 And i <= dhz - 2 And lu(1) = 0 And lu(2) = 0 And lu(3) = 0 Then
               If ew(i - 4, j) = "" And ew(i + 1, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i - 2, j) = "" And ew(i - 3, j) = "" Then
                     lu(4) = 10
                  End If
               End If
        End If
        '///////////////////////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then '����һ��
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = "" And ew(i, j + 3) = "" Then
                 lu(5) = 10
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 And lu(5) = 0 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = "" And ew(i, j + 2) = "" Then
                     lu(6) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 3 And j <= dhz - 3 And lu(5) = 0 And lu(6) = 0 Then
               If ew(i, j - 3) = "" And ew(i, j + 2) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j - 2) = "" And ew(i, j + 1) = "" Then
                     lu(7) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 4 And j <= dhz - 2 And lu(5) = 0 And lu(6) = 0 And lu(7) = 0 Then
               If ew(i, j - 4) = "" And ew(i, j + 1) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j - 2) = "" And ew(i, j - 3) = "" Then
                     lu(8) = 10
                  End If
               End If
        End If
        '////////////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then  '����бһ��
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = "" Then
                 lu(9) = 10
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 And lu(9) = 0 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = "" Then
                     lu(10) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 3 And j <= dhz - 3 And lu(9) = 0 And lu(10) = 0 Then
               If ew(i - 3, j - 3) = "" And ew(i + 2, j + 2) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i - 2, j - 2) = "" And ew(i + 1, j + 1) = "" Then
                     lu(11) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 4 And i <= dhz - 2 And j >= 4 And j <= dhz - 2 And lu(9) = 0 And lu(10) = 0 And lu(11) = 0 Then
               If ew(i - 4, j - 4) = "" And ew(i + 1, j + 1) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i - 2, j - 2) = "" And ew(i - 3, j - 3) = "" Then
                     lu(12) = 10
                  End If
               End If
        End If
        '/////////////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then  '����бһ��
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = "" Then
                 lu(13) = 10
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 And lu(13) = 0 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = "" Then
                     lu(14) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 2 And j <= dhz - 4 And lu(13) = 0 And lu(14) = 0 Then
               If ew(i - 3, j + 3) = "" And ew(i + 2, j - 2) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i - 2, j + 2) = "" And ew(i + 1, j - 1) = "" Then
                     lu(15) = 10
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 4 And i <= dhz - 2 And j >= 1 And j <= dhz - 5 And lu(13) = 0 And lu(14) = 0 And lu(15) = 0 Then
               If ew(i - 4, j + 4) = "" And ew(i + 1, j - 1) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i - 2, j + 2) = "" And ew(i - 3, j + 3) = "" Then
                     lu(16) = 10
                  End If
               End If
        End If
        '/////////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then '���Ŷ���
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = "" And ew(i + 3, j) = "" Then
                 lu(17) = 100
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And lu(17) = 0 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = m And ew(i + 2, j) = "" Then
                     lu(18) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And lu(17) = 0 And lu(18) = 0 Then
               If ew(i - 3, j) = "" And ew(i + 2, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i - 2, j) = "" And ew(i + 1, j) = m Then
                     lu(19) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = m And ew(i + 3, j) = "" Then
                 lu(20) = 150
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And lu(20) = 0 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = "" And ew(i + 2, j) = m Then
                     lu(21) = 150
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = "" And ew(i + 3, j) = m Then
                 lu(22) = 180
              End If
           End If
        End If
        '/////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then '���Ŷ���
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = "" And ew(i, j + 3) = "" Then
                 lu(23) = 100
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 And lu(23) = 0 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = m And ew(i, j + 2) = "" Then
                     lu(24) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 3 And j <= dhz - 3 And lu(23) = 0 And lu(24) = 0 Then
               If ew(i, j - 3) = "" And ew(i, j + 2) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j - 2) = "" And ew(i, j + 1) = m Then
                     lu(25) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = m And ew(i, j + 3) = "" Then
                 lu(26) = 150
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 And lu(26) = 0 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = "" And ew(i, j + 2) = m Then
                     lu(27) = 150
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = "" And ew(i, j + 3) = m Then
                 lu(28) = 180
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then '���϶���
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = "" Then
                 lu(29) = 100
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 And lu(29) = 0 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" Then
                     lu(30) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 3 And j <= dhz - 3 And lu(29) = 0 And lu(30) = 0 Then
               If ew(i - 3, j - 3) = "" And ew(i + 2, j + 2) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i - 2, j - 2) = "" And ew(i + 1, j + 1) = m Then
                     lu(31) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = "" Then
                 lu(32) = 150
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 And lu(32) = 0 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m Then
                     lu(33) = 150
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = m Then
                 lu(34) = 180
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////
       If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then '���¶���
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = "" Then
                 lu(35) = 100
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 And lu(35) = 0 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" Then
                     lu(36) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 2 And j <= dhz - 4 And lu(35) = 0 And lu(36) = 0 Then
               If ew(i - 3, j + 3) = "" And ew(i + 2, j - 2) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i - 2, j + 2) = "" And ew(i + 1, j - 1) = m Then
                     lu(37) = 100
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = "" Then
                 lu(38) = 150
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 And lu(38) = 0 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m Then
                     lu(39) = 150
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = m Then
                 lu(40) = 180
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then '��������
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = m And ew(i + 3, j) = "" Then
                 lu(41) = 1500
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And lu(41) = 0 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = m And ew(i + 2, j) = m Then
                     lu(42) = 1500
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = m And ew(i + 3, j) = m Then
                 lu(43) = 1000
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = "" And ew(i + 3, j) = m Then
                 lu(44) = 1000
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then '��������
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = m And ew(i, j + 3) = "" Then
                 lu(45) = 1500
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 And lu(45) = 0 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = m And ew(i, j + 2) = m Then
                     lu(46) = 1500
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = m And ew(i, j + 3) = m Then
                 lu(47) = 1000
              End If
           End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = "" And ew(i, j + 3) = m Then
                 lu(48) = 1000
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////
       If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then '��������
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = "" Then
                 lu(49) = 1500
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 And lu(49) = 0 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m Then
                     lu(50) = 1500
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = m Then
                 lu(51) = 1000
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = m Then
                 lu(52) = 1000
              End If
           End If
        End If
        '///////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then '��������
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = "" Then
                 lu(53) = 1500
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 And lu(53) = 0 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m Then
                     lu(54) = 1500
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = m Then
                 lu(55) = 1000
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = m Then
                 lu(56) = 1000
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then    '�����Ŀ�����������
           If ew(i - 1, j) = "" And ew(i + 4, j) = n Then
              If ew(i + 1, j) = m And ew(i + 2, j) = m And ew(i + 3, j) = m Then
                 lu(57) = 2000
              End If
           End If
           If ew(i - 1, j) = n And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = m And ew(i + 3, j) = m Then
                 lu(58) = 2000
              End If
           End If
        End If
        If ew(i, j) = m And i <= dhz - 5 Then
           If ew(i + 4, j) = m And ew(i + 2, j) = m Then
              If ew(i + 1, j) = m And ew(i + 3, j) = "" Then
                 lu(59) = 1500
              End If
              If ew(i + 1, j) = "" And ew(i + 3, j) = m Then
                 lu(60) = 1500
              End If
           End If
           If ew(i + 4, j) = m And ew(i + 2, j) = "" And ew(i + 1, j) = m And ew(i + 3, j) = m Then
              lu(61) = 2000
           End If
        End If
        '//////////////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = n Then
              If ew(i, j + 1) = m And ew(i, j + 2) = m And ew(i, j + 3) = m Then
                 lu(62) = 2000
              End If
           End If
           If ew(i, j - 1) = n And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = m And ew(i, j + 3) = m Then
                 lu(63) = 2000
              End If
           End If
        End If
        If ew(i, j) = m And j <= dhz - 5 Then
           If ew(i, j + 4) = m And ew(i, j + 2) = m Then
              If ew(i, j + 1) = m And ew(i, j + 3) = "" Then
                 lu(64) = 1500
              End If
              If ew(i, j + 1) = "" And ew(i, j + 3) = m Then
                 lu(65) = 1500
              End If
           End If
           If ew(i, j + 4) = m And ew(i, j + 2) = "" And ew(i, j + 1) = m And ew(i, j + 3) = m Then
              lu(66) = 2000
           End If
        End If
        '/////////////////////////////////////////////////////////////
        If ew(i, j) = m And i <= dhz - 5 And j <= dhz - 5 And i >= 1 And j >= 1 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = n Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = m Then
                 lu(67) = 2000
              End If
           End If
           If ew(i - 1, j - 1) = n And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = m Then
                 lu(68) = 2000
              End If
           End If
        End If
        If ew(i, j) = m And i <= dhz - 5 And j <= dhz - 5 Then
           If ew(i + 4, j + 4) = m And ew(i + 2, j + 2) = m Then
              If ew(i + 1, j + 1) = m And ew(i + 3, j + 3) = "" Then
                 lu(69) = 1500
              End If
              If ew(i + 1, j + 1) = "" And ew(i + 3, j + 3) = m Then
                 lu(70) = 1500
              End If
           End If
           If ew(i + 4, j + 4) = m And ew(i + 2, j + 2) = "" And ew(i + 1, j + 1) = m And ew(i + 3, j + 3) = m Then
              lu(71) = 2000
           End If
        End If
        '////////////////////////////////////////////////////////////
        If ew(i, j) = m And i <= dhz - 5 And j <= dhz - 2 And i >= 1 And j >= 4 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = n Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = m Then
                 lu(72) = 2000
              End If
           End If
           If ew(i - 1, j + 1) = n And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = m Then
                 lu(73) = 2000
              End If
           End If
        End If
        If ew(i, j) = m And i <= dhz - 5 And j >= 4 Then
           If ew(i + 4, j - 4) = m And ew(i + 2, j - 2) = m Then
              If ew(i + 1, j - 1) = m And ew(i + 3, j - 3) = "" Then
                 lu(74) = 1500
              End If
              If ew(i + 1, j - 1) = "" And ew(i + 3, j - 3) = m Then
                 lu(75) = 1500
              End If
           End If
           If ew(i + 4, j - 4) = m And ew(i + 2, j - 2) = "" And ew(i + 1, j - 1) = m And ew(i + 3, j - 3) = m Then
              lu(76) = 2000
           End If
        End If
        '/////////////////////////////////////////////////////////////
        '//////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then    '�����Ŀ���������
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = m And ew(i + 3, j) = m Then
                 lu(77) = 10000
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = m And ew(i, j + 3) = m Then
                 lu(78) = 10000
              End If
           End If
        End If
        '////////////////////////////////////////////////////////////
        If ew(i, j) = m And i <= dhz - 5 And j <= dhz - 5 And i >= 1 And j >= 1 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = m Then
                 lu(79) = 10000
              End If
           End If
        End If
        '///////////////////////////////////////////////////////////
        If ew(i, j) = m And i <= dhz - 5 And j <= dhz - 2 And i >= 1 And j >= 4 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = m Then
                 lu(80) = 10000
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////
        If ew(i, j) = n And i >= 1 And i <= dhz - 5 Then       '�Է��Ŀ�����
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = n And ew(i + 2, j) = n And ew(i + 3, j) = n Then
                 lu(81) = -100000
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////////
        If ew(i, j) = n And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = n And ew(i, j + 2) = n And ew(i, j + 3) = n Then
                 lu(82) = -100000
              End If
           End If
        End If
        '////////////////////////////////////////////////////////////
        If ew(i, j) = n And i <= dhz - 5 And j <= dhz - 5 And i >= 1 And j >= 1 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = n And ew(i + 2, j + 2) = n And ew(i + 3, j + 3) = n Then
                 lu(83) = -100000
              End If
           End If
        End If
        '///////////////////////////////////////////////////////////
        If ew(i, j) = n And i <= dhz - 5 And j <= dhz - 2 And i >= 1 And j >= 4 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = n And ew(i + 2, j - 2) = n And ew(i + 3, j - 3) = n Then
                 lu(84) = -100000
              End If
           End If
        End If
        '////////////////////////////////////////////////////////////
        
        '/////////////////////////////////////////////////////////////
        For l = 1 To 84
            vz(j) = vz(j) + lu(l)
        Next l
    Next j
    For h = 0 To dhz - 1
        vh(i) = vh(i) + vz(h)
    Next h
Next i
For f = 0 To dhz - 1
    su = su + vh(f)
Next f
If cs = 0 Then
   cs = 1
   estimate = su - estimate(ew(), Not qw)
   cs = 0
ElseIf cs = 1 Then
       estimate = su
End If
End Function

Public Function ab!(ByVal dep%, ByVal pass!)
If dep = 1 Then
   Dim qe As Boolean
   qe = True
   ab = estimate(wez(), qe) - estimate(wez(), Not qe)
Else
   Dim qp%(24, 24), vl!, at%, d%, good!, bad!, w%, ds$(1 To 5), h%, z%
   If dep Mod 2 = 0 Then
      qe = True
   ElseIf edp Mod 2 = 1 Then
          qe = False
   End If
   Call choice(wez(), ds(), qe)
   For k = 1 To 3
       Call jstr(ds(k), h, z)
       qp(h, z) = 1
   Next k
   bad = -10000000
   good = 10000000
   Do
     at = 0
     For i = 0 To dhz - 1
         For j = 0 To dhz - 1
             If qp(i, j) = 1 And dep Mod 2 = 1 Then
                wez(i, j) = "����"
                qp(i, j) = 3
                at = 1
                If bad = -10000000 Then
                   vl = ab(dep - 1, -10000000)
                Else
                    vl = ab(dep - 1, bad)
                End If
                wez(i, j) = ""
                If vl >= bad Then
                   bad = vl
                End If
                If vl > pass Then
                   ab = 10000000
                   Exit Function
                End If
                Exit For
             ElseIf qp(i, j) = 1 And dep Mod 2 = 0 Then
                    wez(i, j) = "����"
                    qp(i, j) = 3
                    at = 1
                    If good = 10000000 Then
                       vl = ab(dep - 1, 10000000)
                    Else
                        vl = ab(dep - 1, good)
                    End If
                    wez(i, j) = ""
                    If vl <= good Then
                       good = vl
                    End If
                    If vl < pass Then
                       ab = -10000000
                       Exit Function
                    End If
                    Exit For
             End If
         Next j
         If at = 1 Then Exit For
     Next i
     d = 0
     For i = 0 To dhz - 1
         For j = 0 To dhz - 1
             If qp(i, j) = 3 Then
                d = d + 1
             End If
         Next j
     Next i
   Loop Until d = 3
   If dep Mod 2 = 0 Then
      ab = good
   ElseIf dep Mod 2 = 1 Then
          ab = bad
   End If
End If
End Function

Public Sub fk(ByVal wq As Boolean)
Dim fh%, fz%, zh%, zz%
If wq = True Then
   FillStyle = 0
   FillColor = ys2
   Call jstr(hz(bsh), fh, fz)
   Picture1.Line (fh * 10 + 2.5, fz * 10 + 7.5)-(fh * 10 + 7.5, fz * 10 + 2.5), ys2, BF
   If bsb >= 1 Then
      Call jstr(bz(bsb), zh, zz)
      For i = 0 To 100
          Picture1.Circle (zh * 10 + 5, zz * 10 + 5), i / 20, ys2
      Next i
   End If
ElseIf wq = False Then
       FillStyle = 0
       FillColor = ys1
       Call jstr(bz(bsb), fh, fz)
       Picture1.Line (fh * 10 + 2.5, fz * 10 + 7.5)-(fh * 10 + 7.5, fz * 10 + 2.5), ys1, BF
       If bsh >= 1 Then
          Call jstr(hz(bsh), zh, zz)
          For i = 0 To 100
              Picture1.Circle (zh * 10 + 5, zz * 10 + 5), i / 20, ys1
          Next i
       End If
End If
End Sub

Public Function two(ew$(), ByVal q As Boolean, ByRef kd$()) As Boolean
Dim m As String * 2, jk%, k(24, 24, 1 To 4) As Boolean
two = False
If q = True Then
   m = "����"
Else
   m = "����"
End If
For i = 0 To 24
    For j = 0 To 24
        For n = 1 To 4
            k(i, j, n) = False
        Next n
    Next j
Next i
jk = 0
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then '���Ŷ���
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = "" And ew(i + 3, j) = "" Then
                 k(i + 2, j, 1) = True: k(i + 3, j, 1) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = m And ew(i + 2, j) = "" Then
                     k(i + 2, j, 1) = True: k(i - 1, j, 1) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 Then
               If ew(i - 3, j) = "" And ew(i + 2, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i - 2, j) = "" And ew(i + 1, j) = m Then
                     k(i - 2, j, 1) = True: k(i - 1, j, 1) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = m And ew(i + 3, j) = "" Then
                 k(i + 1, j, 1) = True: k(i + 3, j, 1) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 Then
               If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                  If ew(i - 1, j) = "" And ew(i + 1, j) = "" And ew(i + 2, j) = m Then
                     k(i - 1, j, 1) = True: k(i + 1, j, 1) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = "" And ew(i + 3, j) = m Then
                 k(i + 2, j, 1) = True: k(i + 1, j, 1) = True
              End If
           End If
        End If
        '/////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then '���Ŷ���
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = "" And ew(i, j + 3) = "" Then
                 k(i, j + 2, 2) = True: k(i, j + 3, 2) = True
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = m And ew(i, j + 2) = "" Then
                     k(i, j + 2, 2) = True: k(i, j - 1, 2) = True
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 3 And j <= dhz - 3 Then
               If ew(i, j - 3) = "" And ew(i, j + 2) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j - 2) = "" And ew(i, j + 1) = m Then
                     k(i, j - 2, 2) = True: k(i, j - 1, 2) = True
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = m And ew(i, j + 3) = "" Then
                 k(i, j + 1, 2) = True: k(i, j + 3, 2) = True
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 Then
               If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = "" And ew(i, j + 2) = m Then
                     k(i, j - 1, 2) = True: k(i, j + 1, 2) = True
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = "" And ew(i, j + 3) = m Then
                 k(i, j + 2, 2) = True: k(i, j + 1, 2) = True
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then '���϶���
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = "" Then
                 k(i + 2, j + 2, 3) = True: k(i + 3, j + 3, 3) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" Then
                     k(i - 1, j - 1, 3) = True: k(i + 2, j + 2, 3) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 3 And j <= dhz - 3 Then
               If ew(i - 3, j - 3) = "" And ew(i + 2, j + 2) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i - 2, j - 2) = "" And ew(i + 1, j + 1) = m Then
                     k(i - 1, j - 1, 3) = True: k(i - 2, j - 2, 3) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = "" Then
                 k(i + 1, j + 1, 3) = True: k(i + 3, j + 3, 3) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 Then
               If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m Then
                     k(i - 1, j - 1, 3) = True: k(i + 1, j + 1, 3) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = m Then
                 k(i + 1, j + 1, 3) = True: k(i + 2, j + 2, 3) = True
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////
       If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then '���¶���
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = "" Then
                 k(i + 2, j - 2, 4) = True: k(i + 3, j - 3, 4) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" Then
                     k(i - 1, j + 1, 4) = True: k(i + 2, j - 2, 4) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 3 And i <= dhz - 3 And j >= 2 And j <= dhz - 4 Then
               If ew(i - 3, j + 3) = "" And ew(i + 2, j - 2) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i - 2, j + 2) = "" And ew(i + 1, j - 1) = m Then
                     k(i - 1, j + 1, 4) = True: k(i - 2, j + 2, 4) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = "" Then
                 k(i + 1, j - 1, 4) = True: k(i + 3, j - 3, 4) = True
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 Then
               If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m Then
                     k(i - 1, j + 1, 4) = True: k(i + 1, j - 1, 4) = True
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = m Then
                 k(i + 1, j - 1, 4) = True: k(i + 2, j - 2, 4) = True
              End If
           End If
        End If
        For l = 0 To 24
            For p = 0 To 24
                For n = 1 To 4
                    If k(l, p, n) = True Then
                       two = True
                       jk = jk + 1
                       ReDim Preserve kd(jk) As String
                       kd(jk) = bstr(l, p) & "," & n & "," & i & j
                       k(l, p, n) = False
                    End If
                Next n
            Next p
        Next l
    Next j
Next i
End Function

Public Function san(ew$(), ByVal r As Boolean, ByRef kd$()) As Boolean
Dim m As String * 2, n As String * 2, lk(24, 24, 1 To 4) As Boolean, ks%
san = False
If r = True Then
   m = "����"
   n = "����"
Else
   m = "����"
   n = "����"
End If
For i = 0 To 24
    For j = 0 To 24
        For k = 1 To 4
            lk(i, j, k) = False
        Next k
    Next j
Next i
ks = 0
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then '��������
           If ew(i - 1, j) = "" Or ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = m And ew(i + 3, j) = "" Then
                 If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
                    If i >= 2 Then
                       If ew(i - 2, j) <> m Then
                          lk(i + 3, j, 1) = True: lk(i + 4, j, 1) = True: lk(i - 1, j, 1) = True: thtw = True
                       End If
                    Else
                        lk(i + 3, j, 1) = True: lk(i + 4, j, 1) = True: lk(i - 1, j, 1) = True: thtw = True
                    End If
                 End If
                 If ew(i - 1, j) = "" And ew(i + 4, j) = n Then
                    If i >= 2 Then
                       If ew(i - 2, j) <> m Then
                          lk(i + 3, j, 1) = True: lk(i - 1, j, 1) = True
                       End If
                    Else
                        lk(i + 3, j, 1) = True: lk(i - 1, j, 1) = True
                    End If
                 End If
                 If ew(i - 1, j) = n And ew(i + 4, j) = "" Then
                    lk(i + 3, j, 1) = True: lk(i + 4, j, 1) = True: thtw = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 Then
           If ew(i - 2, j) = "" Or ew(i + 3, j) = "" Then
              If ew(i - 1, j) = "" And ew(i + 1, j) = m And ew(i + 2, j) = m Then
                 If ew(i - 2, j) = "" And ew(i + 3, j) = "" Then
                    If i <= dhz - 5 Then
                       If ew(i + 4, j) <> m Then
                          lk(i - 1, j, 1) = True: lk(i - 2, j, 1) = True: lk(i + 3, j, 1) = True: thtw = True
                       End If
                    Else
                        lk(i - 1, j, 1) = True: lk(i - 2, j, 1) = True: lk(i + 3, j, 1) = True: thtw = True
                    End If
                 End If
                 If ew(i - 2, j) = "" And ew(i + 3, j) = n Then
                    lk(i - 1, j, 1) = True: lk(i - 2, j, 1) = True: thtw = True
                 End If
                 If ew(i - 2, j) = n And ew(i + 3, j) = "" Then
                    If i <= dhz - 5 Then
                       If ew(i + 4, j) <> m Then
                          lk(i - 1, j, 1) = True: lk(i + 3, j, 1) = True
                       End If
                    Else
                        lk(i - 1, j, 1) = True: lk(i + 3, j, 1) = True
                    End If
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" Or ew(i + 4, j) = "" Then
              If ew(i + 1, j) = "" And ew(i + 2, j) = m And ew(i + 3, j) = m Then
                 If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
                    lk(i - 1, j, 1) = True: lk(i + 4, j, 1) = True: lk(i + 1, j, 1) = True: oot = True
                 End If
                 If ew(i - 1, j) = "" And ew(i + 4, j) = n Then
                    lk(i - 1, j, 1) = True: lk(i + 1, j, 1) = True: oot = True
                 End If
                 If ew(i - 1, j) = n And ew(i + 4, j) = "" Then
                    lk(i + 4, j, 1) = True: lk(i + 1, j, 1) = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 Then
           If ew(i - 1, j) = "" Or ew(i + 4, j) = "" Then
              If ew(i + 1, j) = m And ew(i + 2, j) = "" And ew(i + 3, j) = m Then
                 If ew(i - 1, j) = "" And ew(i + 4, j) = "" Then
                    lk(i + 2, j, 1) = True: lk(i + 4, j, 1) = True: lk(i - 1, j, 1) = True: oot = True
                 End If
                 If ew(i - 1, j) = "" And ew(i + 4, j) = n Then
                    lk(i + 2, j, 1) = True: lk(i - 1, j, 1) = True
                 End If
                 If ew(i - 1, j) = n And ew(i + 4, j) = "" Then
                    lk(i + 2, j, 1) = True: lk(i + 4, j, 1) = True: oot = True
                 End If
              End If
           End If
        End If
        '/////////////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then '��������
           If ew(i, j - 1) = "" Or ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = m And ew(i, j + 3) = "" Then
                 If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
                    If j >= 2 Then
                       If ew(i, j - 2) <> m Then
                          lk(i, j + 3, 2) = True: lk(i, j + 4, 2) = True: lk(i, j - 1, 2) = True: thtw = True
                       End If
                    Else
                        lk(i, j + 3, 2) = True: lk(i, j + 4, 2) = True: lk(i, j - 1, 2) = True: thtw = True
                    End If
                 End If
                 If ew(i, j - 1) = "" And ew(i, j + 4) = n Then
                    If j >= 2 Then
                       If ew(i, j - 2) <> m Then
                          lk(i, j + 3, 2) = True: lk(i, j - 1, 2) = True
                       End If
                    Else
                        lk(i, j + 3, 2) = True: lk(i, j - 1, 2) = True
                    End If
                 End If
                 If ew(i, j - 1) = n And ew(i, j + 4) = "" Then
                    lk(i, j + 3, 2) = True: lk(i, j + 4, 2) = True: thtw = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And j >= 2 And j <= dhz - 4 Then
               If ew(i, j - 2) = "" Or ew(i, j + 3) = "" Then
                  If ew(i, j - 1) = "" And ew(i, j + 1) = m And ew(i, j + 2) = m Then
                     If ew(i, j - 2) = "" And ew(i, j + 3) = "" Then
                        If j <= dhz - 5 Then
                           If ew(i, j + 4) <> m Then
                              lk(i, j + 3, 2) = True: lk(i, j - 2, 2) = True: lk(i, j - 1, 2) = True: thtw = True
                           End If
                        Else
                            lk(i, j + 3, 2) = True: lk(i, j - 2, 2) = True: lk(i, j - 1, 2) = True: thtw = True
                        End If
                     End If
                     If ew(i, j - 2) = "" And ew(i, j + 3) = n Then
                        lk(i, j - 2, 2) = True: lk(i, j - 1, 2) = True: thtw = True
                     End If
                     If ew(i, j - 2) = n And ew(i, j + 3) = "" Then
                        If j <= dhz - 5 Then
                           If ew(i, j + 4) <> m Then
                              lk(i, j + 3, 2) = True: lk(i, j - 1, 2) = True
                           End If
                        Else
                            lk(i, j + 3, 2) = True: lk(i, j - 1, 2) = True
                        End If
                     End If
                  End If
               End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" Or ew(i, j + 4) = "" Then
              If ew(i, j + 1) = "" And ew(i, j + 2) = m And ew(i, j + 3) = m Then
                 If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
                    lk(i, j + 1, 2) = True: lk(i, j + 4, 2) = True: lk(i, j - 1, 2) = True: oot = True
                 End If
                 If ew(i, j - 1) = "" And ew(i, j + 4) = n Then
                    lk(i, j + 1, 2) = True: lk(i, j - 1, 2) = True: oot = True
                 End If
                 If ew(i, j - 1) = n And ew(i, j + 4) = "" Then
                    lk(i, j + 1, 2) = True: lk(i, j + 4, 2) = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And j >= 1 And j <= dhz - 5 Then
           If ew(i, j - 1) = "" Or ew(i, j + 4) = "" Then
              If ew(i, j + 1) = m And ew(i, j + 2) = "" And ew(i, j + 3) = m Then
                 If ew(i, j - 1) = "" And ew(i, j + 4) = "" Then
                    lk(i, j + 2, 2) = True: lk(i, j + 4, 2) = True: lk(i, j - 1, 2) = True: oot = True
                 End If
                 If ew(i, j - 1) = "" And ew(i, j + 4) = n Then
                    lk(i, j + 2, 2) = True: lk(i, j - 1, 2) = True
                 End If
                 If ew(i, j - 1) = n And ew(i, j + 4) = "" Then
                    lk(i, j + 2, 2) = True: lk(i, j + 4, 2) = True: oot = True
                 End If
              End If
           End If
        End If
        '//////////////////////////////////////////////////////////
       If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then '��������
           If ew(i - 1, j - 1) = "" Or ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = "" Then
                 If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
                    If i >= 2 And j >= 2 Then
                       If ew(i - 2, j - 2) <> m Then
                          lk(i + 3, j + 3, 3) = True: lk(i + 4, j + 4, 3) = True: lk(i - 1, j - 1, 3) = True: thtw = True
                       End If
                    Else
                        lk(i + 3, j + 3, 3) = True: lk(i + 4, j + 4, 3) = True: lk(i - 1, j - 1, 3) = True: thtw = True
                    End If
                 End If
                 If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = n Then
                    If i >= 2 And j >= 2 Then
                       If ew(i - 2, j - 2) <> m Then
                          lk(i + 3, j + 3, 3) = True: lk(i - 1, j - 1, 3) = True
                       End If
                    Else
                        lk(i + 3, j + 3, 3) = True: lk(i - 1, j - 1, 3) = True
                    End If
                 End If
                 If ew(i - 1, j - 1) = n And ew(i + 4, j + 4) = "" Then
                    lk(i + 3, j + 3, 3) = True: lk(i + 4, j + 4, 3) = True: thtw = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 2 And j <= dhz - 4 Then
               If ew(i - 2, j - 2) = "" Or ew(i + 3, j + 3) = "" Then
                  If ew(i - 1, j - 1) = "" And ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = m Then
                     If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = "" Then
                        If i <= dhz - 5 And j <= dhz - 5 Then
                           If ew(i + 4, j + 4) <> m Then
                              lk(i + 3, j + 3, 3) = True: lk(i - 2, j - 2, 3) = True: lk(i - 1, j - 1, 3) = True: thtw = True
                           End If
                        Else
                            lk(i + 3, j + 3, 3) = True: lk(i - 2, j - 2, 3) = True: lk(i - 1, j - 1, 3) = True: thtw = True
                        End If
                     End If
                     If ew(i - 2, j - 2) = "" And ew(i + 3, j + 3) = n Then
                        lk(i - 2, j - 2, 3) = True: lk(i - 1, j - 1, 3) = True: thtw = True
                     End If
                     If ew(i - 2, j - 2) = n And ew(i + 3, j + 3) = "" Then
                        If j <= dhz - 5 And i <= dhz - 5 Then
                           If ew(i + 4, j + 4) <> m Then
                              lk(i + 3, j + 3, 3) = True: lk(i - 1, j - 1, 3) = True
                           End If
                        Else
                            lk(i + 3, j + 3, 3) = True: lk(i - 1, j - 1, 3) = True
                        End If
                     End If
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" Or ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = m Then
                If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
                   lk(i + 1, j + 1, 3) = True: lk(i + 4, j + 4, 3) = True: lk(i - 1, j - 1, 3) = True: oot = True
                End If
                If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = n Then
                   lk(i + 1, j + 1, 3) = True: lk(i - 1, j - 1, 3) = True: oot = True
                End If
                If ew(i - 1, j - 1) = n And ew(i + 4, j + 4) = "" Then
                   lk(i + 1, j + 1, 3) = True: lk(i + 4, j + 4, 3) = True
                End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 1 And j <= dhz - 5 Then
           If ew(i - 1, j - 1) = "" Or ew(i + 4, j + 4) = "" Then
              If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = m Then
                 If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = "" Then
                    lk(i + 2, j + 2, 3) = True: lk(i + 4, j + 4, 3) = True: lk(i - 1, j - 1, 3) = True: oot = True
                 End If
                 If ew(i - 1, j - 1) = "" And ew(i + 4, j + 4) = n Then
                    lk(i + 2, j + 2, 3) = True: lk(i - 1, j - 1, 3) = True
                 End If
                 If ew(i - 1, j - 1) = n And ew(i + 4, j + 4) = "" Then
                    lk(i + 2, j + 2, 3) = True: lk(i + 4, j + 4, 3) = True: oot = True
                 End If
              End If
           End If
        End If
        '///////////////////////////////////////////////////////////
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then '��������
           If ew(i - 1, j + 1) = "" Or ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = "" Then
                 If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
                    If i >= 2 And j <= dhz - 3 Then
                       If ew(i - 2, j + 2) <> m Then
                          lk(i + 3, j - 3, 4) = True: lk(i + 4, j - 4, 4) = True: lk(i - 1, j + 1, 4) = True: thtw = True
                       End If
                    Else
                        lk(i + 3, j - 3, 4) = True: lk(i + 4, j - 4, 4) = True: lk(i - 1, j + 1, 4) = True: thtw = True
                    End If
                 End If
                 If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = n Then
                    If i >= 2 And j <= dhz - 3 Then
                       If ew(i - 2, j + 2) <> m Then
                          lk(i + 3, j - 3, 4) = True: lk(i - 1, j + 1, 4) = True
                       End If
                    Else
                        lk(i + 3, j - 3, 4) = True: lk(i - 1, j + 1, 4) = True
                    End If
                 End If
                 If ew(i - 1, j + 1) = n And ew(i + 4, j - 4) = "" Then
                    lk(i + 3, j - 3, 4) = True: lk(i + 4, j - 4, 4) = True: thtw = True
                 End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 2 And i <= dhz - 4 And j >= 3 And j <= dhz - 3 Then
               If ew(i - 2, j + 2) = "" Or ew(i + 3, j - 3) = "" Then
                  If ew(i - 1, j + 1) = "" And ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = m Then
                     If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = "" Then
                        If j >= 4 And i <= dhz - 5 Then
                           If ew(i + 4, j - 4) <> m Then
                              lk(i + 3, j - 3, 4) = True: lk(i - 2, j + 2, 4) = True: lk(i - 1, j + 1, 4) = True: thtw = True
                           End If
                        Else
                            lk(i + 3, j - 3, 4) = True: lk(i - 2, j + 2, 4) = True: lk(i - 1, j + 1, 4) = True: thtw = True
                        End If
                     End If
                     If ew(i - 2, j + 2) = "" And ew(i + 3, j - 3) = n Then
                        lk(i - 2, j + 2, 4) = True: lk(i - 1, j + 1, 4) = True: thtw = True
                     End If
                     If ew(i - 2, j + 2) = n And ew(i + 3, j - 3) = "" Then
                        If j >= 4 And i <= dhz - 5 Then
                           If ew(i + 4, j - 4) <> m Then
                              lk(i + 3, j - 3, 4) = True: lk(i - 1, j + 1, 4) = True
                           End If
                        Else
                            lk(i + 3, j - 3, 4) = True: lk(i - 1, j + 1, 4) = True
                        End If
                     End If
                  End If
               End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" Or ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = m Then
                 If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
                   lk(i - 1, j + 1, 4) = True: lk(i + 4, j - 4, 4) = True: lk(i + 1, j - 1, 4) = True: oot = True
                End If
                If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = n Then
                   lk(i - 1, j + 1, 4) = True: lk(i + 1, j - 1, 4) = True: oot = True
                End If
                If ew(i - 1, j + 1) = n And ew(i + 4, j - 4) = "" Then
                   lk(i + 1, j - 1, 4) = True: lk(i + 4, j - 4, 4) = True
                End If
              End If
           End If
        End If
        If ew(i, j) = m And i >= 1 And i <= dhz - 5 And j >= 4 And j <= dhz - 2 Then
           If ew(i - 1, j + 1) = "" Or ew(i + 4, j - 4) = "" Then
              If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = m Then
                 If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = "" Then
                    lk(i + 2, j - 2, 4) = True: lk(i + 4, j - 4, 4) = True: lk(i - 1, j + 1, 4) = True: oot = True
                 End If
                 If ew(i - 1, j + 1) = "" And ew(i + 4, j - 4) = n Then
                    lk(i + 2, j - 2, 4) = True: lk(i - 1, j + 1, 4) = True
                 End If
                 If ew(i - 1, j + 1) = n And ew(i + 4, j - 4) = "" Then
                    lk(i + 2, j - 2, 4) = True: lk(i + 4, j - 4, 4) = True: oot = True
                 End If
              End If
           End If
        End If
        For l = 0 To 24
            For p = 0 To 24
                For k = 1 To 4
                    If lk(l, p, k) = True Then
                       san = True
                       ks = ks + 1
                       ReDim Preserve kd(1 To ks) As String
                       kd(ks) = bstr(l, p) & "," & k & "," & i & j
                       lk(l, p, k) = False
                    End If
                Next k
            Next p
        Next l
    Next j
Next i
End Function

Public Function er(ByRef rh%, ByRef rz%, ByVal we As Boolean) As Boolean
Dim sf(1 To 4) As Boolean, h1%, z1%, yt$, yu$
er = False
If we = True Then
   yt = "����"
   yu = "����"
Else
    yt = "����"
    yu = "����"
End If
                    For i = 1 To 4
                        sf(i) = False
                    Next i
                    For i = 0 To dhz - 1
                        For j = 0 To dhz - 1
                            If i >= 1 And i <= dhz - 2 Then
                               If wz(i, j) = wz(i + 1, j) And wz(i, j) = yt Then
                                  If wz(i - 1, j) = "" And wz(i + 2, j) = "" Then
                                     sf(1) = True
                                     h1 = i: z1 = j
                                     Exit For
                                  End If
                                End If
                            End If
                            If j >= 1 And j <= dhz - 3 Then
                               If wz(i, j) = wz(i, j + 1) And wz(i, j) = yt Then
                                  If wz(i, j - 1) = "" And wz(i, j + 2) = "" Then
                                     sf(2) = True
                                     h1 = i: z1 = j
                                     Exit For
                                  End If
                               End If
                            End If
                            If i >= 1 And j >= 1 And i <= dhz - 3 And j <= dhz - 3 Then
                               If wz(i, j) = wz(i + 1, j + 1) And wz(i, j) = yt Then
                                  If wz(i - 1, j - 1) = "" And wz(i + 2, j + 2) = "" Then
                                     sf(3) = True
                                     h1 = i: z1 = j
                                     Exit For
                                  End If
                               End If
                            End If
                            If i <= dhz - 3 And j >= 2 And i >= 1 And j <= dhz - 2 Then
                               If wz(i, j) = wz(i + 1, j - 1) And wz(i, j) = yt Then
                                  If wz(i - 1, j + 1) = "" And wz(i + 2, j - 2) = "" Then
                                     sf(4) = True
                                     h1 = i: z1 = j
                                     Exit For
                                  End If
                               End If
                            End If
                        Next j
                        If sf(1) = True Then Exit For
                        If sf(2) = True Then Exit For
                        If sf(3) = True Then Exit For
                        If sf(4) = True Then Exit For
                    Next i
                    For i = 1 To 4
                        If sf(i) = True Then
                           er = True
                           Dim jzd1!, jzd2!
                           Select Case True
                                      Case sf(1)
                                           wz(h1 - 1, z1) = yu
                                           jzd1 = estimate(wz(), sf(1))
                                           wz(h1 - 1, z1) = ""
                                           wz(h1 + 2, z1) = yu
                                           jzd2 = estimate(wz(), sf(1))
                                           wz(h1 + 2, z1) = ""
                                           If jzd1 >= jzd2 Then
                                              rh = h1 - 1: rz = z1
                                           Else
                                               rh = h1 + 2: rz = z1
                                           End If
                                      Case sf(2)
                                           wz(h1, z1 - 1) = yu
                                           jzd1 = estimate(wz(), sf(2))
                                           wz(h1, z1 - 1) = ""
                                           wz(h1, z1 + 2) = yu
                                           jzd2 = estimate(wz(), sf(2))
                                            wz(h1, z1 + 2) = ""
                                           If jzd1 >= jzd2 Then
                                              rh = h1: rz = z1 - 1
                                           Else
                                               rh = h1: rz = z1 + 2
                                           End If
                                      Case sf(3)
                                           wz(h1 - 1, z1 - 1) = yu
                                           jzd1 = estimate(wz(), sf(3))
                                           wz(h1 - 1, z1 - 1) = ""
                                           wz(h1 + 2, z1 + 2) = yu
                                           jzd2 = estimate(wz(), sf(3))
                                            wz(h1 + 2, z1 + 2) = ""
                                           If jzd1 >= jzd2 Then
                                              rh = h1 - 1: rz = z1 - 1
                                           Else
                                               rh = h1 + 2: rz = z1 + 2
                                           End If
                                      Case sf(4)
                                           wz(h1 - 1, z1 + 1) = yu
                                           jzd1 = estimate(wz(), sf(4))
                                           wz(h1 - 1, z1 + 1) = ""
                                           wz(h1 + 2, z1 - 2) = yu
                                           jzd2 = estimate(wz(), sf(4))
                                            wz(h1 + 2, z1 - 2) = ""
                                           If jzd1 >= jzd2 Then
                                              rh = h1 - 1: rz = z1 + 1
                                           Else
                                               rh = h1 + 2: rz = z1 - 2
                                           End If
                           End Select
                           Exit For
                        End If
                    Next i
End Function

Public Function rsan(ew$(), ByVal Y As Boolean, ByRef kd$()) As Boolean
Dim m As String * 2, lk(24, 24, 1 To 4) As Boolean, rk%
rsan = False
If Y = True Then
   m = "����"
Else
   m = "����"
End If
For i = 0 To 24
    For j = 0 To 24
        For k = 1 To 4
            lk(i, j, k) = False
        Next k
    Next j
Next i
rk = 0
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If ew(i, j) = m And i <= dhz - 5 Then
           If ew(i + 1, j) = m And ew(i + 2, j) = "" And ew(i + 3, j) = "" And ew(i + 4, j) = m Then
              lk(i + 2, j, 1) = True: lk(i + 3, j, 1) = True: tto = True
           End If
        End If
        If ew(i, j) = m And i <= dhz - 5 Then
           If ew(i + 1, j) = "" And ew(i + 2, j) = m And ew(i + 3, j) = "" And ew(i + 4, j) = m Then
              lk(i + 1, j, 1) = True: lk(i + 3, j, 1) = True: gyg = True
           End If
        End If
        If ew(i, j) = m And i <= dhz - 5 Then
           If ew(i + 1, j) = "" And ew(i + 2, j) = "" And ew(i + 3, j) = m And ew(i + 4, j) = m Then
              lk(i + 2, j, 1) = True: lk(i + 1, j, 1) = True: tto = True
           End If
        End If
        '////////////////////////////////////////
        If ew(i, j) = m And j <= dhz - 5 Then
           If ew(i, j + 1) = m And ew(i, j + 2) = "" And ew(i, j + 3) = "" And ew(i, j + 4) = m Then
              lk(i, j + 2, 2) = True: lk(i, j + 3, 2) = True: tto = True
           End If
        End If
        If ew(i, j) = m And j <= dhz - 5 Then
           If ew(i, j + 1) = "" And ew(i, j + 2) = m And ew(i, j + 3) = "" And ew(i, j + 4) = m Then
              lk(i, j + 1, 2) = True: lk(i, j + 3, 2) = True: gyg = True
           End If
        End If
        If ew(i, j) = m And j <= dhz - 5 Then
           If ew(i, j + 1) = "" And ew(i, j + 2) = "" And ew(i, j + 3) = m And ew(i, j + 4) = m Then
              lk(i, j + 2, 2) = True: lk(i, j + 1, 2) = True: tto = True
           End If
        End If
        '///////////////////////////////////////////////
        If ew(i, j) = m And j <= dhz - 5 And i <= dhz - 5 Then
           If ew(i + 1, j + 1) = m And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = "" And ew(i + 4, j + 4) = m Then
              lk(i + 2, j + 2, 3) = True: lk(i + 3, j + 3, 3) = True: tto = True
           End If
        End If
        If ew(i, j) = m And j <= dhz - 5 And i <= dhz - 5 Then
           If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = m And ew(i + 3, j + 3) = "" And ew(i + 4, j + 4) = m Then
              lk(i + 1, j + 1, 3) = True: lk(i + 3, j + 3, 3) = True: gyg = True
           End If
        End If
        If ew(i, j) = m And j <= dhz - 5 And i <= dhz - 5 Then
           If ew(i + 1, j + 1) = "" And ew(i + 2, j + 2) = "" And ew(i + 3, j + 3) = m And ew(i + 4, j + 4) = m Then
              lk(i + 2, j + 2, 3) = True: lk(i + 1, j + 1, 3) = True: tto = True
           End If
        End If
        '////////////////////////////////////////////////////
        If ew(i, j) = m And j >= 4 And i <= dhz - 5 Then
           If ew(i + 1, j - 1) = m And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = "" And ew(i + 4, j - 4) = m Then
              lk(i + 2, j - 2, 4) = True: lk(i + 3, j - 3, 4) = True: tto = True
           End If
        End If
        If ew(i, j) = m And j >= 4 And i <= dhz - 5 Then
           If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = m And ew(i + 3, j - 3) = "" And ew(i + 4, j - 4) = m Then
              lk(i + 1, j - 1, 4) = True: lk(i + 3, j - 3, 4) = True: gyg = True
           End If
        End If
        If ew(i, j) = m And j >= 4 And i <= dhz - 5 Then
           If ew(i + 1, j - 1) = "" And ew(i + 2, j - 2) = "" And ew(i + 3, j - 3) = m And ew(i + 4, j - 4) = m Then
              lk(i + 2, j - 2, 4) = True: lk(i + 1, j - 1, 4) = True: tto = True
           End If
        End If
        '/////////////////////////////////////////////////////
        For l = 0 To 24
            For p = 0 To 24
                For k = 1 To 4
                    If lk(l, p, k) = True Then
                       rsan = True
                       rk = rk + 1
                       ReDim Preserve kd(rk) As String
                       kd(rk) = bstr(l, p) & "," & k & "," & i & j
                       lk(l, p, k) = False
                    End If
                Next k
            Next p
        Next l
    Next j
Next i
End Function

Public Sub chqp()
If Picture1.Height > Picture2.Height And Picture1.Width > Picture2.Width Then
   If h1 = 0 Then
      ma = -pt2 * HS.Value
   Else
       ma = -pt2 * HS.Value * h1
   End If
   If v1 = 0 Then
      mb = -(Picture1.Height - Picture2.Height - VS.Value) * pt1
   Else
       mb = -(Picture1.Height - Picture2.Height - VS.Value * v1) * pt1
   End If
   If h1 = 0 Then
      mc = (Picture1.Width - HS.Value) * pt2
   Else
       mc = (Picture1.Width - HS.Value * h1) * pt2
   End If
   If v1 = 0 Then
      mdd = (Picture2.Height + VS.Value) * pt1
   Else
       mdd = (Picture2.Height + VS.Value * v1) * pt1
   End If
   Picture1.Scale (ma, mdd)-(mc, mb)
ElseIf Picture1.Height > Picture2.Height And Picture1.Width <= Picture2.Width Then
       Picture1.Width = Picture2.Width
       If v1 = 0 Then
          mdd = (Picture2.Height + VS.Value) * pt1
       Else
           mdd = (Picture2.Height + VS.Value * v1) * pt1
       End If
       If v1 = 0 Then
          mb = -(Picture1.Height - Picture2.Height - VS.Value) * pt1
       Else
           mb = -(Picture1.Height - Picture2.Height - VS.Value * v1) * pt1
       End If
       Picture1.Scale (0, mdd)-(dhz * 10, mb)
ElseIf Picture1.Height <= Picture2.Height And Picture1.Width > Picture2.Width Then
       Picture1.Height = Picture2.Height
       If h1 = 0 Then
          ma = -pt2 * HS.Value
       Else
           ma = -pt2 * HS.Value * h1
       End If
       If h1 = 0 Then
          mc = (Picture1.Width - HS.Value) * pt2
       Else
           mc = (Picture1.Width - HS.Value * h1) * pt2
       End If
       Picture1.Scale (ma, dhz * 10)-(mc, 0)
End If
Call hqp
If fzqz.Checked = False Then
For i = 0 To 24
    For j = 0 To 24
        If wz(i, j) = "����" Then
           For m = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), m / 20, ys1
           Next m
        End If
        If wz(i, j) = "����" Then
           For l = 1 To 100
               Picture1.Circle (i * 10 + 5, j * 10 + 5), l / 20, ys2
           Next l
        End If
    Next j
Next i
If bsb >= 1 Or bsh >= 1 Then
   If slz = True Then
      Call fk(Not slz)
   ElseIf slz = False Then
          Call fk(Not slz)
   End If
End If
ElseIf fzqz.Checked = True Then
       Dim qh%, qz%
       If bsh >= 1 Then
          For i = 1 To bsh
              Call jstr(hz(i), qh, qz)
              Imah(i).Left = qh * 10
              Imah(i).Top = (qz + 1) * 10
          Next i
       End If
       If bsb >= 1 Then
          For i = 1 To bsb
              Call jstr(bz(i), qh, qz)
              Imab(i).Left = qh * 10
              Imab(i).Top = (qz + 1) * 10
          Next i
       End If
End If
End Sub

Public Sub kjbj()
If md = 1 Then
   hfys.Caption = "�ڷ���ɫ": bfys.Caption = "�׷���ɫ"
   hy.Caption = "�ڷ���ɫ": bye.Caption = "�׷���ɫ"
   Option1.Caption = "�ڷ�����": Option2.Caption = "�׷�����"
   Lal1.Caption = "�ڷ�": Lal2.Caption = "�׷�"
   Comby.Enabled = True: Comhy.Enabled = True
   If fzqz.Checked = False Then
      Comby.Visible = True: Comhy.Visible = True
      ys.Enabled = True: tcys.Enabled = True
   End If
   Comjl.Visible = False
   Textip.Visible = False
   Comlj.Visible = False
   Frame1.Visible = True
   lisdh.Visible = False
   Textdh.Visible = False
   Picture2.Enabled = True
   Labip.Caption = "": zykj.Enabled = True
   szqp.Enabled = True: qpsz.Enabled = True: jsxz.Enabled = True
   scqp.Enabled = True: bfys = True: hfys.Enabled = True
   szqp.Enabled = True: qpsz.Enabled = True
   yxxz.Enabled = True: ckbc.Enabled = True
ElseIf md = 2 Then
       hfys.Caption = "������ɫ": bfys.Caption = dl.mz & "��ɫ"
       hy.Caption = "������ɫ": bye.Caption = dl.mz & "��ɫ"
       Option1.Caption = "��������": Option2.Caption = dl.mz & "����"
       If fzqz.Checked = False Then
          Comby.Visible = True: Comhy.Visible = True
          ys.Enabled = True: tcys.Enabled = True
       End If
       Lal1.Caption = "����": Lal2.Caption = dl.mz
       Comjl.Visible = False
       Textip.Visible = False
       Comlj.Visible = False
       Frame1.Visible = True
       lisdh.Visible = False
       Textdh.Visible = False
       Picture2.Enabled = True: hfys.Enabled = True
       Labip.Caption = "": zykj.Enabled = True
       szqp.Enabled = True: qpsz.Enabled = True: jsxz.Enabled = True
       scqp.Enabled = True: bfys.Enabled = True
       szqp.Enabled = True: qpsz.Enabled = True
       yxxz.Enabled = True: ckbc.Enabled = True
ElseIf md = 3 Then
       hfys.Caption = dl.mz & "��ɫ": bfys.Caption = "����" & "��ɫ"
       hy.Caption = dl.mz & "��ɫ": bye.Caption = "����" & "��ɫ"
       Comby.Visible = False: Comhy.Visible = False
       Lal1.Caption = dl.mz: Lal2.Caption = "����"
       Labip.Caption = "": hfys.Enabled = False
       Comjl.Visible = True
       Textip.Visible = True
       Comlj.Visible = True
       Picture2.Enabled = False
       Frame1.Visible = False
       lisdh.Visible = True
       Textdh.Visible = True
       hy.Enabled = False
       lisdh = "��Ϣ���տ�"
       Textdh.Text = "�ڴ�������Ϣ�����س�����"
       Textip.Text = "���������IP��ַ"
       laizou = False: bfys.Enabled = False
       zykj.Enabled = False: scqp.Enabled = False
       szqp.Enabled = False: qpsz.Enabled = False
       yxxz.Enabled = False: ckbc.Enabled = False
End If
End Sub

Public Sub choice(wez$(), ByRef sh$(), ByVal rt As Boolean)
Dim kh!(24, 24), zd!, zdt$, m%, n%, X$
If rt = True Then
   X = "����"
Else
   X = "����"
End If
For i = 0 To dhz - 1
    For j = 0 To dhz - 1
        If wez(i, j) <> "" Then
           For a = i - 2 To i + 2
               For b = j - 2 To j + 2
                   If a >= 0 And a <= dhz - 1 And b >= 0 And b <= dhz - 1 Then
                      If wez(a, b) = "" And kh(a, b) = 0 Then
                         wez(a, b) = X
                         kh(a, b) = estimate(wez(), rt)
                         wez(a, b) = ""
                      End If
                   End If
               Next b
           Next a
        End If
    Next j
Next i
For k = 1 To 3
    zd = 0
    For i = 0 To dhz - 1
        For j = 0 To dhz - 1
            If kh(i, j) >= zd Then
               zd = kh(i, j)
               zdt = bstr(i, j)
            End If
        Next j
    Next i
    Call jstr(zdt, m, n)
    kh(m, n) = 0
    sh(k) = zdt
Next k
End Sub

Public Function aa!(ByVal dep%, ByVal pass!)
If dep = 1 Then
   Dim qe As Boolean
   qe = True
   aa = estimate(wez(), qe)
Else
   Dim qp(24, 24) As Boolean, vl!, good!, bad!
   For i = 0 To dhz - 1
       For j = 0 To dhz - 1
           qp(i, j) = False
           If wez(i, j) <> "" Then
              For a = i - 2 To i + 2
                  For b = j - 2 To j + 2
                      If a >= 0 And a <= dhz - 1 And b >= 0 And b <= dhz - 1 Then
                         If wez(a, b) = "" And qp(a, b) <> True Then
                            qp(a, b) = True
                         End If
                      End If
                  Next b
              Next a
           End If
       Next j
   Next i
   bad = -10000000
   good = 10000000
   For i = 0 To dhz - 1
       For j = 0 To dhz - 1
           If qp(i, j) = True And dep Mod 2 = 1 Then
                wez(i, j) = "����"
                If bad = -10000000 Then
                   vl = aa(dep - 1, -10000000)
                Else
                    vl = aa(dep - 1, bad)
                End If
                wez(i, j) = ""
                If vl >= bad Then
                   bad = vl
                End If
                If vl > pass Then
                   aa = 10000000
                   Exit Function
                End If
                Exit For
           ElseIf qp(i, j) = True And dep Mod 2 = 0 Then
                    wez(i, j) = "����"
                    If good = 10000000 Then
                       vl = aa(dep - 1, 10000000)
                    Else
                        vl = aa(dep - 1, good)
                    End If
                    wez(i, j) = ""
                    If vl <= good Then
                       good = vl
                    End If
                    If vl < pass Then
                       aa = -10000000
                       Exit Function
                    End If
                    Exit For
           End If
       Next j
   Next i
   If dep Mod 2 = 0 Then
      aa = good
   ElseIf dep Mod 2 = 1 Then
          aa = bad
   End If
End If
End Function

Private Sub hy_Click()
Call Comhy_Click
End Sub

Private Sub bye_Click()
Call Comby_Click
End Sub

Public Function compare(zh$(), fa$(), cb As Boolean) As String
Dim zf As Boolean, fn!, fe!, fh%, fz%      'cbΪ�����������ڹ�ͬ��Ӯ��cbΪ�������Լ�����ѵ��
fn = -1000000
compare = ""
If UBound(zh()) = UBound(fa()) Then
   zf = True
   For m = 1 To UBound(zh())
       For n = 1 To UBound(fa())
           If m = n Then
              If zh(m) <> fa(n) Then
                 zf = False
                 Exit For
              End If
           End If
       Next n
       If zf = False Then
          Exit For
       End If
   Next m
   If m > UBound(zh()) And n > UBound(fa()) Then
      zf = True:
      For i = 1 To UBound(zh())
          For j = i To UBound(zh())
              If i < j Then
                 sa = Split(zh(i), ",")
                 sz = Split(zh(j), ",")
                 If ggz = True And gyg = True Then
                        If sa(0) = sz(0) Then
                           compare = compare & sa(0)
                        End If
                 ElseIf ssz = True And thtw = True Then
                        If sa(0) = sz(0) Then
                           compare = compare & sa(0)
                        End If
                 Else
                 
                 If sa(0) = sz(0) And sa(1) <> sz(1) Then
                    If cb = True Then
                       compare = compare & sa(0)
                    Else
                        For k = 1 To UBound(zh())
                            sc = Split(zh(k), ",")
                            If sa(2) = sc(2) And sa(1) = sc(1) Then
                               Call jstr(sc(0), fh, fz)
                               wz(fh, fz) = "����"
                               fe = estimate(wz(), Not cb)
                               wz(fh, fz) = ""
                               If fn < fe Then
                                  fn = fe
                                  compare = bstr(fh, fz)
                               End If
                            End If
                            If sz(2) = sc(2) And sz(1) = sc(1) Then
                               Call jstr(sc(0), fh, fz)
                               wz(fh, fz) = "����"
                               fe = estimate(wz(), Not cb)
                               wz(fh, fz) = ""
                               If fn < fe Then
                                  fn = fe
                                  compare = bstr(fh, fz)
                               End If
                            End If
                        Next k
                    End If
                 End If
                 End If
              End If
          Next j
      Next i
   Else
       zf = False
   End If
Else
    zf = False
End If
If zf = False Then
   For i = 1 To UBound(zh())
       For j = 1 To UBound(fa())
           sa = Split(zh(i), ",")
           sz = Split(fa(j), ",")
           If rrz = True And tto = True And oot = True Then
                 
                 If sa(0) = sz(0) Then
                    compare = compare & sa(0)
                 End If
           Else
           
           If sa(0) = sz(0) And sa(1) <> sz(1) Then
              If cb = True Then
                 compare = compare & sa(0)
              Else
                  For k = 1 To UBound(zh())
                      sc = Split(zh(k), ",")
                      If sa(2) = sc(2) And sa(1) = sc(1) Then
                         Call jstr(sc(0), fh, fz)
                         wz(fh, fz) = "����"
                         fe = estimate(wz(), Not cb)
                         wz(fh, fz) = ""
                         If fn < fe Then
                            fn = fe
                            compare = bstr(fh, fz)
                         End If
                      End If
                  Next k
                  For p = 1 To UBound(fa())
                      sc = Split(fa(p), ",")
                      If sz(2) = sc(2) And sz(1) = sc(1) Then
                         Call jstr(sc(0), fh, fz)
                         wz(fh, fz) = "����"
                         fe = estimate(wz(), Not cb)
                         wz(fh, fz) = ""
                         If fn < fe Then
                            fn = fe
                            compare = bstr(fh, fz)
                         End If
                      End If
                  Next p
              End If
              End If
           End If
       Next j
   Next i
End If
End Function


Public Sub qxzt()
If qxts.Checked = True Or (ssjs.Checked = True Or sijs.Checked = True Or cljs.Checked = True) Then
Dim zm$, ko$(), xh%, xz%, qx$, gf As Boolean, rh As Boolean, sh As Boolean, _
th As Boolean, jl As Boolean, dgh$(), dkh$(), dfh$(), gd$
For i = 1 To 2
 rh = False: th = False: jl = False: sh = False
If i = 1 Then
   If md = 1 Then
      zm = "�ڷ�"
   ElseIf md = 2 Then
          zm = "����"
   ElseIf md = 3 Then
          zm = Lal1.Caption
   End If
   gf = True
ElseIf i = 2 Then
    If md = 1 Then
       zm = "�׷�"
    ElseIf md = 2 Then
           zm = dl.mz
    ElseIf md = 3 Then
           zm = Lal2.Caption
    End If
    gf = False
End If
If four(wz(), xh, xz, gf) = 2 And qxts.Checked = True Then
   qx = zm & "�ɻ�����"
   jl = True
End If
If rsan(wz(), gf, dgh()) = True Then
   rh = True
   ggz = True
   gd = compare(dgh(), dgh(), True)
   ggz = False
   If gd <> "" Then
      If (Option1.Value = True And sijs.Checked = True And i = 1 And md <> 3) Or (Option2.Value = True And sijs.Checked = True And i = 2 And md <> 3) Or (md = 3 And zhc = "zhu" And sijs.Checked = True) Then
         fourj = gd & zm
      End If
      qx = zm & "��������"
      jl = True
   End If
End If
If san(wz(), gf, dfh()) = True Then
   sh = True
   ssz = True
   gd = compare(dfh(), dfh(), True)
   ssz = False
   If gd <> "" Then
      If (Option1.Value = True And sijs.Checked = True And i = 1 And md <> 3) Or (Option2.Value = True And sijs.Checked = True And i = 2 And md <> 3) Or (md = 3 And zhc = "zhu" And sijs.Checked = True) Then
            If fourj <> "" Then
               fourj = Left(fourj, Len(fourj) - Len(zm)) & gd & zm
            Else
                fourj = gd & zm
            End If
      End If
      qx = zm & "��������"
      jl = True
   End If
End If
If rh = True And sh = True Then
   rrz = True
   gd = compare(dfh(), dgh(), True)
   rrz = False
   If gd <> "" Then
      If (Option1.Value = True And sijs.Checked = True And i = 1 And md <> 3) Or (Option2.Value = True And sijs.Checked = True And i = 2 And md <> 3) Or (md = 3 And zhc = "zhu" And sijs.Checked = True) Then
         If fourj <> "" Then
               fourj = Left(fourj, Len(fourj) - Len(zm)) & gd & zm
            Else
                fourj = gd & zm
            End If
      End If
      qx = zm & "��������"
      jl = True
   End If
End If
If jl = False And two(wz(), gf, dkh()) = True Then th = True
If th = True And rh = True And jl = False And qxts.Checked = True Then
   If compare(dkh(), dgh(), True) <> "" Then
      qx = zm & "��������"
      jl = True
   End If
End If
If th = True And sh = True And jl = False And qxts.Checked = True Then
   If compare(dkh(), dfh(), True) <> "" Then
      qx = zm & "��������"
      jl = True
   End If
End If
If four(wz(), xh, xz, gf) = 1 And jl = False And qxts.Checked = True Then
       qx = zm & "�ɳ�����"
       jl = True
ElseIf three(wz(), xh, xz, gf) = True And jl = False And qxts.Checked = True Then
       qx = zm & "����������"
       jl = True
ElseIf twoone(wz(), xh, xz, gf) = True And jl = False And qxts.Checked = True Then
       qx = zm & "����������"
       jl = True
End If
If th = True Then
   gd = compare(dkh(), dkh(), True)
   If gd <> "" Then
      If (Option1.Value = True And i = 1 And ssjs.Checked = True And md <> 3) Or (Option2.Value = True And i = 2 And ssjs.Checked = True And md <> 3) Or (md = 3 And i = 2 And zhc = "zhu" And ssjs.Checked = True) Then
         threej = gd & zm
      End If
      qx = zm & "��������"
      jl = True
   End If
End If
oot = False: thtw = False: gyg = False: tto = False
If two(wz(), gf, ko()) = True And jl = False And qxts.Checked = True Then
       qx = zm & "�ɻ����"
End If
If gf = True And qxts.Checked = True Then
   qxh = qx
   qx = ""
   If qxh <> "" And qxb = "" Then
      Labelts.Caption = qxh
   End If
   If qxh = "" And qxb <> "" Then
      Labelts.Caption = qxb
   End If
   If qxh <> "" And qxb <> "" Then
      Labelts.Caption = qxh & "����" & qxb
   End If
   If qxh = "" And qxb = "" Then
      Labelts.Caption = ""
   End If
ElseIf gf = False And qxts.Checked = True Then
    qxb = qx
    qx = ""
    If qxh <> "" And qxb = "" Then
       Labelts.Caption = qxh
    End If
    If qxh = "" And qxb <> "" Then
       Labelts.Caption = qxb
    End If
    If qxh <> "" And qxb <> "" Then
       Labelts.Caption = qxh & "����" & qxb
    End If
    If qxh = "" And qxb = "" Then
      Labelts.Caption = ""
    End If
End If
Next i
End If
If qxts.Checked = False Then
    Labelts.Caption = ""
End If
End Sub

Public Sub hqp() '������
Dim h%
Picture1.Cls
For i = 0 To dhz - 1
    Picture1.Line (5 + 10 * i, 5)-(5 + 10 * i, dhz * 10 - 5)
    Picture1.Line (5, 5 + 10 * i)-(dhz * 10 - 5, 5 + 10 * i)
Next i
Select Case dhz
        Case 9, 11
             h = 3
        Case 13, 15, 17
             h = 4
        Case 19, 21
             h = 5
        Case 23, 25
             h = 6
End Select
For i = 1 To 50
    Picture1.Circle (dhz * 5, dhz * 5), i / 50, 0
    Picture1.Circle (h * 10 - 5, h * 10 - 5), i / 50, 0
    Picture1.Circle (dhz * 10 - h * 10 + 5, dhz * 10 - h * 10 + 5), i / 50, 0
    Picture1.Circle (h * 10 - 5, dhz * 10 - h * 10 + 5), i / 50, 0
    Picture1.Circle (dhz * 10 - h * 10 + 5, h * 10 - 5), i / 50, 0
Next i
End Sub
Private Sub yxsm_Click()
ave = 2
Formsm.Show 1
End Sub
Private Sub ztl_Click()
ztl.Checked = Not ztl.Checked
If ztl.Checked = True Then
   Picsta.Visible = True
   qxts.Checked = True
   qxts.Enabled = True
Else
    Picsta.Visible = False
    qxts.Checked = False
    qxts.Enabled = False
End If
End Sub

Private Sub zz_Click(index As Integer)
Dim h1$, h2$, b$
Select Case index
       Case 1
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr((dhz - 1) / 2, (dhz - 1) / 2)
          h2 = bstr((dhz - 1) / 2, (dhz - 1) / 2 + 2)
          b = bstr((dhz - 1) / 2, (dhz - 1) / 2 + 1)
       Case 2
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 1), ((dhz - 1) / 2 + 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 3
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 + 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 2), ((dhz - 1) / 2 + 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 4
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 1), ((dhz - 1) / 2 + 1))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 5
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 2), ((dhz - 1) / 2 + 1))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 6
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 1), ((dhz - 1) / 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 7
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 2), ((dhz - 1) / 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 8
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 - 1))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 9
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 1), ((dhz - 1) / 2 - 1))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 10
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 2) * 10 + 5, ((dhz - 1) / 2 - 1) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 2), ((dhz - 1) / 2 - 1))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 11
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 - 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
       Case 12
          For i = 1 To 100
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2) * 10 + 5), i / 20, RGB(0, 0, 0)
              Picture1.Circle (((dhz - 1) / 2) * 10 + 5, ((dhz - 1) / 2 + 1) * 10 + 5), i / 20, RGB(255, 255, 255)
              Picture1.Circle (((dhz - 1) / 2 + 1) * 10 + 5, ((dhz - 1) / 2 - 2) * 10 + 5), i / 20, RGB(0, 0, 0)
          Next i
          h1 = bstr(((dhz - 1) / 2), ((dhz - 1) / 2))
          h2 = bstr(((dhz - 1) / 2 + 1), ((dhz - 1) / 2 - 2))
          b = bstr(((dhz - 1) / 2), ((dhz - 1) / 2 + 1))
End Select
tcys.Enabled = False
ys.Enabled = False
Comhy.Visible = False
Comby.Visible = False
dzms.Enabled = False
qpsz.Enabled = False
szqp.Enabled = False
zykj.Enabled = False
If ys1 = 0 And ys2 = 0 Then
   ys1 = RGB(0, 0, 0): ys2 = RGB(255, 255, 255)
End If
Dim shu%, m1%, n1%, m2%, n2%, m3%, n3%
If md = 2 Then
   shu = MsgBox("��ѡ���ķ��������(Ҳ�����ߴ˾ְ���)��" & vbCrLf & "���ǡ�Ϊ���ԣ�����Ϊ���", 36, "��ѡ��")
   Picture1.Cls
   Call hqp
   If shu = vbYes Then
      hz(1) = b: bz(1) = h1: bz(2) = h2
      Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
      wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
      Call hqz(m1, n1, 1, False)
      Call hqz(m2, n2, 2, False)
      Call hqz(m3, n3, 1, True)
      Call jstr(autolz(wz()), m1, n1)
      Labelzb.Caption = Chr(65 + m1) & n1 + 1 & " " & "���ӵ㣺" & Chr(65 + m1) & n1 + 1
      lzd = "���ӵ㣺" & Chr(65 + m1) & n1 + 1
      Call hqz(m1, n1, 2, True)
      wz(m1, n1) = "����"
      bsh = 1
      bsh = bsh + 1
      hz(bsh) = bstr(m1, n1)
      bsb = 2
      Labelbsb.Caption = "��" & bsb & "��"
      Labelbsh.Caption = "��" & bsh & "��"
      slz = True
      If fzqz.Checked = False Then
         Call fk(slz)
      End If
      slz = False
      Timerb.Enabled = True
      Option2.Value = True
      Frame1.Caption = "��Ϸ�У�����ѡ��"
      Frame1.Enabled = False
      Randomize
      Dim rand%
      rand = Rnd * 3
      If rand < 1 Then
         rand = 2
      ElseIf rand < 2 Then
             rand = 3
      ElseIf rand < 3 Then
             rand = 6
      End If
      If rand = 2 Then
         Call ssjs_Click
         Call sijs_Click
      End If
      If rand = 3 Then
         Call cljs_Click
         Call ssjs_Click
      End If
      If rand = 6 Then
         Call sijs_Click
         Call cljs_Click
      End If
      Call qxzt
   ElseIf shu = vbNo Then
          hz(1) = h1: hz(2) = h2: bz(1) = b
          Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
          wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
          Call hqz(m1, n1, 1, True)
          Call hqz(m2, n2, 2, True)
          Call hqz(m3, n3, 1, False)
          slz = False
          bsh = 2: bsb = 1
          Labelbsh.Caption = "��" & bsh & "��"
          Labelbsb.Caption = "��" & bsb & "��"
          Timerb = True
          Option1.Value = True
          Frame1.Caption = "��Ϸ�У�����ѡ��"
          Frame1.Enabled = False
   End If
ElseIf md = 1 Then
       Picture1.Cls
       Call hqp
       hz(1) = h1: hz(2) = h2: bz(1) = b
       Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
       wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
       Call hqz(m1, n1, 1, True)
       Call hqz(m2, n2, 2, True)
       Call hqz(m3, n3, 1, False)
       slz = False
       bsh = 2: bsb = 1
       Labelbsh.Caption = "��" & bsh & "��"
       Labelbsb.Caption = "��" & bsb & "��"
       Timerb = True
       Option1.Value = True
       Frame1.Caption = "��Ϸ�У�����ѡ��"
       Frame1.Enabled = False
ElseIf md = 3 Then
       Dim zc$
       zc = Right(zhc, 3)
       If zc = "zhu" Then
          zhc = "zhu"
          shu = MsgBox("��ѡ���ķ��������(Ҳ�����ߴ˾ְ���)��" & vbCrLf & "���ǡ�Ϊ" & Lal1 & "������Ϊ" & Lal2, 36, "��ѡ��")
          Picture1.Cls
          Call hqp
          If shu = vbYes Then
             hz(1) = b: bz(1) = h1: bz(2) = h2
             Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
             wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
             Call hqz(m1, n1, 1, False)
             Call hqz(m2, n2, 2, False)
             Call hqz(m3, n3, 1, True)
             bsh = 1: bsb = 2
             Labelbsb.Caption = "��" & bsb & "��"
             Labelbsh.Caption = "��" & bsh & "��"
             slz = True
             If fzqz.Checked = False Then
                Call fk(Not slz)
             End If
             Timerh.Enabled = True
             Option2.Value = True
             Frame1.Caption = "��Ϸ�У�����ѡ��"
             Frame1.Enabled = False
             Call qxzt
             Picture2.Enabled = True
             If Win(1).State = sckConnected Then
                Win(1).SendData ("Y" & index & "zz")
             End If
          ElseIf shu = vbNo Then
                 hz(1) = h1: hz(2) = h2: bz(1) = b
                 Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                 wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                 Call hqz(m1, n1, 1, True)
                 Call hqz(m2, n2, 2, True)
                 Call hqz(m3, n3, 1, False)
                 slz = False
                 If fzqz.Checked = False Then
                    Call fk(Not slz)
                 End If
                 bsh = 2: bsb = 1
                 Labelbsh.Caption = "��" & bsh & "��"
                 Labelbsb.Caption = "��" & bsb & "��"
                 Timerb = True
                 Option1.Value = True
                 Frame1.Caption = "��Ϸ�У�����ѡ��"
                 Frame1.Enabled = False
                 Picture2.Enabled = False
                 If Win(1).State = sckConnected Then
                    Win(1).SendData ("N" & index & "zz")
                 End If
          End If
       ElseIf zc = "bei" Then
              zc = Left(zhc, 1)
              zhc = "bei"
              Picture1.Cls
              Call hqp
              If zc = "Y" Then
                 hz(1) = h1: hz(2) = h2: bz(1) = b
                 Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                 wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                 Call hqz(m1, n1, 1, True)
                 Call hqz(m2, n2, 2, True)
                 Call hqz(m3, n3, 1, False)
                 slz = False
                 If fzqz.Checked = False Then
                    Call fk(Not slz)
                 End If
                 bsh = 2: bsb = 1
                 Labelbsh.Caption = "��" & bsh & "��"
                 Labelbsb.Caption = "��" & bsb & "��"
                 Timerb = True
                 Option1.Value = True
                 Frame1.Caption = "��Ϸ�У�����ѡ��"
                 Frame1.Enabled = False
                 Picture2.Enabled = False
                 MsgBox Lal2 & "ʹ��ְҵ���֣�" & zz(index).Caption, 0, "��ʾ"
              ElseIf zc = "N" Then
                     hz(1) = b: bz(1) = h1: bz(2) = h2
                     Call jstr(h1, m1, n1): Call jstr(h2, m2, n2): Call jstr(b, m3, n3)
                     wz(m1, n1) = "����": wz(m2, n2) = "����": wz(m3, n3) = "����"
                     Call hqz(m1, n1, 1, False)
                     Call hqz(m2, n2, 2, False)
                     Call hqz(m3, n3, 1, True)
                     bsh = 1: bsb = 2
                     Labelbsb.Caption = "��" & bsb & "��"
                     Labelbsh.Caption = "��" & bsh & "��"
                     slz = True
                     If fzqz.Checked = False Then
                        Call fk(Not slz)
                     End If
                     Timerh.Enabled = True
                     Option2.Value = True
                     Frame1.Caption = "��Ϸ�У�����ѡ��"
                     Frame1.Enabled = False
                     Call qxzt
                     Picture2.Enabled = True
                     MsgBox Lal2 & "ʹ��ְҵ���֣�" & zz(index).Caption, 0, "��ʾ"
              End If
       End If
End If
End Sub

Private Sub hqz(hh%, zz%, sbh%, sl As Boolean)
If fzqz.Checked = True Then
   If sl = True Then
      Imah(sbh).Height = 10: Imah(sbh).Width = 10
      Imah(sbh).Left = hh * 10: Imah(sbh).Top = (zz + 1) * 10
      Imah(sbh).Visible = True
   ElseIf sl = False Then
          Imab(sbh).Height = 10: Imab(sbh).Width = 10
          Imab(sbh).Left = hh * 10: Imab(sbh).Top = (zz + 1) * 10
          Imab(sbh).Visible = True
   End If
ElseIf fzqz.Checked = False Then
       Dim yanse!
       If sl = True Then
          yanse = ys1
       ElseIf sl = False Then
              yanse = ys2
       End If
       For i = 1 To 100
           Picture1.Circle (hh * 10 + 5, zz * 10 + 5), i / 20, yanse
       Next i
End If
End Sub

Private Sub wzqsm()
smq(1) = "��Ϸ������ͼƬΪԭʼ��С�����������������������������Ĺ������鿴ͼƬ����ط���"
smq(2) = "�����������������Ҽ��������˵����ڸü�������ʱ���ܵ�������ж����Ϸ����ͼƬ��"
smq(3) = "���ֻ���Զ������jpg,gif��bmp��ʽ��ͼƬΪ������"
smq(4) = "��������ı������ѡ�����������ͼƬ����ҿ���������һ��ͼƬ�ϵ��Ҽ����Ӷ��趨�Լ�ϲ����ͼƬ��"
smq(5) = "��Ϸ����Ҳ���Զ�����������Լ�ϲ������ɫ����ע�ⲻҪ��������ɫ��ϡ�"
smq(6) = "��Ϸ���������������趨��ɫ��ͬʱ��ť����ɫ����仯��"
smq(7) = "�ڲ��÷�������ʱ��������ɫ��������Լ��趨��ע������������ɫ��Ҫ��ͬ��"
smq(8) = "ʹ�÷�������ʱ��������ɫ���ܸı�"
smq(9) = "�ڷ������ӺͷǷ�������֮�����ʱ�л���"
smq(10) = "�ڼ������׺������Ϸ�У��������Ͻǲ鿴�ķ���������һ����ע�������ģʽ�޴˹��ܡ�"
smq(11) = "���á���ʼ����Ϸ����ȥ��Ϸ���κ�����״������ע��Ҫ�ȱ������ס�"
smq(12) = "�����ϴ���Ϸ���ס��Ǵ��ϴ��Ѿ�������Ӯ����֣���ע�ⲻ���ϴ�δ��ɻ�Ϊ�������֡�"
smq(13) = "������Ϸ����ʱ�������ף��κ�ģʽ�¾��ɡ�"
smq(14) = "�����ļ��ĺ�׺����lsl��Ҳ����ɭ������֣�������Ϸ���ֻ�ܴ򿪴��ָ�ʽ�����ס�"
smq(15) = "����Ϸ�ṩ9�ִ�С�����̣���Ϸǰ��ѡ����һ��������ģʽ�������ӵ�����趨��"
smq(16) = "���˻�ģʽʱ��ѡ������ԽС�����������ٶ�Խ�죬��֮Խ����"
smq(17) = "�ھ�����ģʽ������Ϊ�Ƿ���ʱ�����������ѡ��˭���ߺ󣬴�ʱ�������ü���������ɫ��"
smq(18) = "����ʾ״̬��ʱ����������ʾ��"
smq(19) = "��������ʾʱ����Ϸ���и��ӿ쳩������˵���ο����������顱��"
smq(20) = "�á�������ʾ��ѡ����һЩ����£��Ƿ���ʾ��ұ������ס�"
smq(21) = "����Ϸ˫��������10����10�����ϣ������¿�ʼ����Ϸʱ���ż�¼δ������һ�Ρ�"
smq(22) = "ͳ�Ƶ���ģʽ����ʱ�䣬�����ȣ����Ǽ���ڰ�˫����ƽ������"
smq(23) = "��Ϸ���Ƴ������⣬����Ϸ˫���������ã�˫��ƽ���ܵ����ơ�"
smq(24) = "������ģʽʱ��˫����������ʱ�䲽�������ƣ�Ҳ����ʱȡ�����ơ�"
smq(25) = "����ֻ�����ߵ�һ����Ч������������һ��ֻ��������ȡʤ��"
smq(26) = "��������Ҽ��˵���ȡ��ȫ����Ϸ���ƣ�"
smq(27) = "�ھ�����ģʽʱ��ֻ�������ӵ�һ��ѡ��ְҵ���֡�"
smq(28) = "�ھ�����ģʽʱ�����������ף�Ҳ�ɴ��ϴ���Ϸ���ף���ֻ�������ӵ�һ��ѡ��"
smq(29) = "�ھ�����ģʽʱ�����壬��ʼ����Ϸ����Ҫ��Է�����"
smq(30) = "����ֻ�ܼ���С�ڸ����̵����ף�Ҳ������С���ݣ�"
smq(31) = "���˻�ģʽʱ��������ص����׸õ��Է����ӣ����Ի��Զ���һ�ӣ�"
smq(32) = "�����������ʾ˫�����ӵ���ɫ��"
smq(33) = "���������IP��ַ�����ӵ����������뽨����������ֻ�ð��¡����������������ɡ�"
smq(34) = "�رշ�������Ͽ����ӣ���ֶ������㣬������ʾ���棬Ҳ�����Զ����档"
smq(35) = "��ת����Ϸģʽ����ر���Ϸʱ�������Զ��Ͽ����������ӡ�"
smq(0) = "ע�⣡������Ϸ���������ļ�����һ����Ϊ��zcb.lsn�����ļ������������Һͳ����������Ϣ������ɾ���൱�ڵ�һ�ν������Ϸ��"
smq(37) = "״̬������ʾ�ϴ����ӵ㣬���÷�������ʱ����������λ�ñ�ʶ�����Կɹۿ��˴��ƶ����鿴��"
smq(36) = "��������˵�������½ǵĽ����ΪA1,��������Ϊ��B,C,D,E����������������Ϊ��2��3��4��5�����������Ͻ����������̴�С������"
smq(38) = "���˻�ģʽʱ�������������ӣ����Ի��Զ�Ϊ����������������֣���ҿ�ȡ�����֡�"
smq(39) = "���¿�ʼ��Ϸʱ�����Զ�ȡ���������ƣ�"
smq(40) = "���˻�ģʽ�����嵽�к���ʱ�����Լ��������ӣ����������ٶȻ������"
smq(41) = "���ý��ֻ�Ӵ���Եļ�������������Һ͵��Ե������ٶȶ�����������"
smq(42) = "������ʱ������Ϣ���տ��ϵ��Ҽ������������壡"
End Sub

Private Function jsix(ch%, cz%, kl As Boolean) As Boolean
Dim jls%(1 To 16)
jsix = False
If kl = True Then
   wz(ch, cz) = "����"
ElseIf kl = False Then
       wz(ch, cz) = "����"
End If
For i = 1 To 5
    If cz <= dhz - 5 And cz >= 1 Then '1-3Ϊ���ϣ����Һ������ϵ�һ�ĿűȽ�
       If wz(ch, cz - 1) = wz(ch, cz - 1 + i) And wz(ch, cz - 1 + i) <> "" Then
          jls(1) = jls(1) + 1
       End If
    End If
    If ch <= dhz - 5 And ch >= 1 Then
       If wz(ch - 1, cz) = wz(ch - 1 + i, cz) And wz(ch - 1 + i, cz) <> "" Then
          jls(2) = jls(2) + 1
       End If
    End If
    If ch <= dhz - 5 And cz <= dhz - 5 And ch >= 1 And cz >= 1 Then
       If wz(ch - 1, cz - 1) = wz(ch - 1 + i, cz - 1 + i) And wz(ch - 1 + i, cz - 1 + i) <> "" Then
          jls(3) = jls(3) + 1
       End If
    End If
    If cz >= 4 And cz <= dhz - 2 Then '4-6Ϊ���£�����������µ�һ�ĿűȽ�
       If wz(ch, cz + 1) = wz(ch, cz + 1 - i) And wz(ch, cz + 1 - i) <> "" Then
          jls(4) = jls(4) + 1
       End If
    End If
    If ch >= 4 And ch <= dhz - 2 Then
       If wz(ch + 1, cz) = wz(ch + 1 - i, cz) And wz(ch + 1 - i, cz) <> "" Then
          jls(5) = jls(5) + 1
       End If
    End If
    If ch >= 4 And cz >= 4 And ch <= dhz - 2 And cz <= dhz - 2 Then
       If wz(ch + 1, cz + 1) = wz(ch + 1 - i, cz + 1 - i) And wz(ch + 1 - i, cz + 1 - i) <> "" Then
          jls(6) = jls(6) + 1
       End If
    End If
    If ch >= 4 And cz <= dhz - 5 And ch <= dhz - 2 And cz >= 1 Then '7-8Ϊ��б�Ϻ���б�µ�һ�ĿűȽ�
       If wz(ch + 1, cz - 1) = wz(ch + 1 - i, cz - 1 + i) And wz(ch + 1 - i, cz - 1 + i) <> "" Then
          jls(7) = jls(7) + 1
       End If
    End If
    If ch <= dhz - 5 And cz >= 4 And ch >= 1 And cz <= dhz - 2 Then
       If wz(ch - 1, cz + 1) = wz(ch - 1 + i, cz + 1 - i) And wz(ch - 1 + i, cz + 1 - i) <> "" Then
          jls(8) = jls(8) + 1
        End If
    End If
    If ch >= 2 And ch <= dhz - 4 Then    '13-17Ϊ���ң�����б�ϣ����ϣ�����б�ϣ�����Ķ����űȽ�
       If wz(ch - 2, cz) = wz(ch - 2 + i, cz) And wz(ch - 2 + i, cz) <> "" Then
          jls(9) = jls(9) + 1
       End If
    End If
    If ch >= 2 And cz >= 2 And ch <= dhz - 4 And cz <= dhz - 4 Then
       If wz(ch - 2, cz - 2) = wz(ch - 2 + i, cz - 2 + i) And wz(ch - 2 + i, cz - 2 + i) <> "" Then
          jls(10) = jls(10) + 1
       End If
    End If
    If cz >= 2 And cz <= dhz - 4 Then
       If wz(ch, cz - 2) = wz(ch, cz - 2 + i) And wz(ch, cz - 2 + i) <> "" Then
          jls(11) = jls(11) + 1
       End If
    End If
    If ch <= dhz - 3 And cz >= 2 And ch >= 3 And cz <= dhz - 4 Then
       If wz(ch + 2, cz - 2) = wz(ch + 2 - i, cz - 2 + i) And wz(ch + 2 - i, cz - 2 + i) <> "" Then
          jls(12) = jls(12) + 1
       End If
    End If
    If ch >= 3 And ch <= dhz - 3 Then
       If wz(ch + 2, cz) = wz(ch + 2 - i, cz) And wz(ch + 2 - i, cz) <> "" Then
          jls(13) = jls(13) + 1
       End If
    End If
    If ch >= 3 And cz >= 3 And ch <= dhz - 3 And cz <= dhz - 3 Then '18-20Ϊ����б�£����£�����б�µĶ����Ƚ�
       If wz(ch + 2, cz + 2) = wz(ch + 2 - i, cz + 2 - i) And wz(ch + 2 - i, cz + 2 - i) <> "" Then
          jls(14) = jls(14) + 1
       End If
    End If
    If cz >= 3 And cz <= dhz - 3 Then
       If wz(ch, cz + 2) = wz(ch, cz + 2 - i) And wz(ch, cz + 2 - i) <> "" Then
          jls(15) = jls(15) + 1
       End If
    End If
    If ch >= 2 And cz <= dhz - 3 And ch <= dhz - 4 And cz >= 3 Then
       If wz(ch - 2, cz + 2) = wz(ch - 2 + i, cz + 2 - i) And wz(ch - 2 + i, cz + 2 - i) <> "" Then
          jls(16) = jls(16) + 1
       End If
    End If
Next i
For i = 1 To 16
    If jls(i) = 5 Then
       jsix = True
       Exit For
    End If
Next i
wz(ch, cz) = ""
End Function
