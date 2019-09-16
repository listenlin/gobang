VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Formsj 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8130
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog Com 
      Left            =   3360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保 存 成 文 档"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确 定 退 出"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Framet 
      BackColor       =   &H00C0FFC0&
      Height          =   3735
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   7815
      Begin VB.TextBox Tet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame Framey 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7815
      Begin VB.Frame Frame4 
         BackColor       =   &H000000FF&
         Caption         =   "局域网模式"
         Height          =   3495
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   1995
         Begin VB.Label Labww 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   2640
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labwh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labws 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labwy 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H008080FF&
         Caption         =   "人机模式"
         Height          =   3495
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   1995
         Begin VB.Label Labrw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   2640
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labrh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labrs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labry 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单人模式"
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3585
         Begin VB.Label Labdbw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1920
            TabIndex        =   14
            Top             =   2640
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdbh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1920
            TabIndex        =   13
            Top             =   1920
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdbs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1920
            TabIndex        =   12
            Top             =   1200
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdby 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1920
            TabIndex        =   11
            Top             =   480
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdhw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   1605
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdhh 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1920
            Width           =   1605
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdhs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1605
            WordWrap        =   -1  'True
         End
         Begin VB.Label Labdhy 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1605
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "白方："
            Height          =   180
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   540
         End
         Begin VB.Line Line1 
            X1              =   1800
            X2              =   1800
            Y1              =   120
            Y2              =   3480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "黑方："
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   540
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7435
      MultiRow        =   -1  'True
      TabFixedHeight  =   882
      HotTracking     =   -1  'True
      TabMinWidth     =   2470
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " 原 始 数 据"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " 统 计 数 据"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Formsj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum$
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Com.CancelError = True
On Error GoTo errhandler
Com.ShowSave
Open FileName & ".txt" For Output As #1
     Print #1, sum
Close #1
errhandler:
End Sub

Private Sub Form_Load()
Formsj.Caption = dl.mz & "的游戏对战数据"
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Dim yl As dlm
Open App.Path & "zcb.lsn" For Random As #1 Len = Len(yl)
    For i = 1 To LOF(1) / Len(yl)
        Get #1, i, yl
        If yl.mz = dl.mz Then
           dl = yl
           Exit For
        End If
    Next i
Close #1
Framey.Visible = True
Framet.Visible = False
Tet.Locked = True
Command2.Visible = False
Call xssj(1)
Dim ran!(1 To 8)
Randomize
For i = 1 To 8
    ran(i) = Int(Rnd * (RGB(255, 255, 255) + 1))
Next i
Command1.BackColor = ran(1)
Command2.BackColor = ran(2)
Me.BackColor = ran(3)
Frame2.BackColor = ran(4)
Frame3.BackColor = ran(5)
Frame4.BackColor = ran(6)
Framet.BackColor = ran(7)
Framey.BackColor = ran(8)
End Sub

Private Sub xssj(i As Integer)
If i = 1 Then
Labdhy.Caption = "总共已赢" & dl.drh.win_ & "次" & vbCrLf & "共用" & _
dl.drh.bs_w & "步" & vbCrLf & "共用" & dl.drh.sj_w \ 3600 & "时" _
& (dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600) \ 60 & "分" & _
dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600 - ((dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600) \ 60) * 60 & "秒"
Labdhs.Caption = "总共已输" & dl.drh.fail & "次" & vbCrLf & "共用" & _
dl.drh.bs_f & "步" & vbCrLf & "共用" & dl.drh.sj_f \ 3600 & "时" _
& (dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600) \ 60 & "分" & _
dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600 - ((dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600) \ 60) * 60 & "秒"
Labdhh.Caption = "总共已和棋" & dl.drh.tie & "次" & vbCrLf & "共用" & _
dl.drh.bs_t & "步" & vbCrLf & "共用" & dl.drh.sj_t \ 3600 & "时" _
& (dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600) \ 60 & "分" & _
dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600 - ((dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600) \ 60) * 60 & "秒"
Labdhw.Caption = "总共未完成棋局" & dl.drh.undone & "次" & vbCrLf & "共用" & _
dl.drh.bs_u & "步" & vbCrLf & "共用" & dl.drh.sj_u \ 3600 & "时" _
& (dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600) \ 60 & "分" & _
dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600 - ((dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600) \ 60) * 60 & "秒"
'///////////////////////////////////////////////////////////////////////////
Labdby.Caption = "总共已赢" & dl.drb.win_ & "次" & vbCrLf & "共用" & _
dl.drb.bs_w & "步" & vbCrLf & "共用" & dl.drb.sj_w \ 3600 & "时" _
& (dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600) \ 60 & "分" & _
dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600 - ((dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600) \ 60) * 60 & "秒"
Labdbs.Caption = "总共已输" & dl.drb.fail & "次" & vbCrLf & "共用" & _
dl.drb.bs_f & "步" & vbCrLf & "共用" & dl.drb.sj_f \ 3600 & "时" _
& (dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600) \ 60 & "分" & _
dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600 - ((dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600) \ 60) * 60 & "秒"
Labdbh.Caption = "总共已和棋" & dl.drb.tie & "次" & vbCrLf & "共用" & _
dl.drb.bs_t & "步" & vbCrLf & "共用" & dl.drb.sj_t \ 3600 & "时" _
& (dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600) \ 60 & "分" & _
dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600 - ((dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600) \ 60) * 60 & "秒"
Labdbw.Caption = "总共未完成棋局" & dl.drb.undone & "次" & vbCrLf & "共用" & _
dl.drb.bs_u & "步" & vbCrLf & "共用" & dl.drb.sj_u \ 3600 & "时" _
& (dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600) \ 60 & "分" & _
dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600 - ((dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600) \ 60) * 60 & "秒"
'///////////////////////////////////////////////////////////////////////////
Labr.Caption = dl.mz & "："
Labry.Caption = "总共已赢" & dl.rj.win_ & "次" & vbCrLf & "共用" & _
dl.rj.bs_w & "步" & vbCrLf & "共用" & dl.rj.sj_w \ 3600 & "时" _
& (dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600) \ 60 & "分" & _
dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600 - ((dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600) \ 60) * 60 & "秒"
Labrs.Caption = "总共已输" & dl.rj.fail & "次" & vbCrLf & "共用" & _
dl.rj.bs_f & "步" & vbCrLf & "共用" & dl.rj.sj_f \ 3600 & "时" _
& (dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600) \ 60 & "分" & _
dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600 - ((dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600) \ 60) * 60 & "秒"
Labrh.Caption = "总共已和棋" & dl.rj.tie & "次" & vbCrLf & "共用" & _
dl.rj.bs_t & "步" & vbCrLf & "共用" & dl.rj.sj_t \ 3600 & "时" _
& (dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600) \ 60 & "分" & _
dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600 - ((dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600) \ 60) * 60 & "秒"
Labrw.Caption = "总共未完成棋局" & dl.rj.undone & "次" & vbCrLf & "共用" & _
dl.rj.bs_u & "步" & vbCrLf & "共用" & dl.rj.sj_u \ 3600 & "时" _
& (dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600) \ 60 & "分" & _
dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600 - ((dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600) \ 60) * 60 & "秒"
'/////////////////////////////////////////////////////////////////////////
Labw.Caption = dl.mz & "："
Labwy.Caption = "总共已赢" & dl.wl.win_ & "次" & vbCrLf & "共用" & _
dl.wl.bs_w & "步" & vbCrLf & "共用" & dl.wl.sj_w \ 3600 & "时" _
& (dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600) \ 60 & "分" & _
dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600 - ((dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600) \ 60) * 60 & "秒"
Labws.Caption = "总共已输" & dl.wl.fail & "次" & vbCrLf & "共用" & _
dl.wl.bs_f & "步" & vbCrLf & "共用" & dl.wl.sj_f \ 3600 & "时" _
& (dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600) \ 60 & "分" & _
dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600 - ((dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600) \ 60) * 60 & "秒"
Labwh.Caption = "总共已和棋" & dl.wl.tie & "次" & vbCrLf & "共用" & _
dl.wl.bs_t & "步" & vbCrLf & "共用" & dl.wl.sj_t \ 3600 & "时" _
& (dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600) \ 60 & "分" & _
dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600 - ((dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600) \ 60) * 60 & "秒"
Labww.Caption = "总共未完成棋局" & dl.wl.undone & "次" & vbCrLf & "共用" & _
dl.wl.bs_u & "步" & vbCrLf & "共用" & dl.wl.sj_u \ 3600 & "时" _
& (dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600) \ 60 & "分" & _
dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600 - ((dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600) \ 60) * 60 & "秒"
ElseIf i = 2 Then
       On Error Resume Next
       Dim wsum%, fsum%, tsum%, usum%, sjsum!, sjz!, bssum!, bsz!, bi$, dh$
       dh = "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝" & vbCrLf
       bi = "整体统计：" & vbCrLf
       sum = bi & dh
       wsum = Round((dl.drh.win_ + dl.drb.win_) / 2) + dl.rj.win_ + dl.wl.win_
       fsum = Round((dl.drh.fail + dl.drb.fail) / 2) + dl.rj.fail + dl.wl.fail
       tsum = Round((dl.drh.tie + dl.drb.tie) / 2) + dl.rj.tie + dl.wl.tie
       usum = Round((dl.drh.undone + dl.drb.undone) / 2) + dl.rj.undone + dl.wl.undone
       sjsum = Round((dl.drh.sj_f + dl.drh.sj_w + dl.drh.sj_t + dl.drh.sj_u + dl.drb.sj_f + dl.drb.sj_w _
              + dl.drb.sj_t + dl.drb.sj_u) / 2) + dl.rj.sj_f + dl.rj.sj_w + dl.rj.sj_t + dl.rj.sj_u _
              + dl.wl.sj_f + dl.wl.sj_w + dl.wl.sj_t + dl.wl.sj_u
       bssum = Round((dl.drh.bs_f + dl.drh.bs_w + dl.drh.bs_t + dl.drh.bs_u + dl.drb.bs_f + dl.drb.bs_w _
               + dl.drb.bs_t + dl.drb.bs_u) / 2) + dl.rj.bs_f + dl.rj.bs_w + dl.rj.bs_t + dl.rj.bs_u _
               + dl.wl.bs_f + dl.wl.bs_w + dl.wl.bs_t + dl.wl.bs_u
       sjz = sjsum: bsz = bssum
       bi = dl.mz & "共用时间：" & sjsum \ 3600 & "时" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "分" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "秒" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "共用步数：" & bssum & "步" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "每步所用时间为" & Format(sjsum / bssum, "0.00") & "秒" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "总共赢：输：和棋：未完成 = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "获胜比率=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "失败比率=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "和棋比率=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "未完成比率=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////////
       bi = "单人模式统计：" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = Round((dl.drh.win_ + dl.drb.win_) / 2)
       fsum = Round((dl.drh.fail + dl.drb.fail) / 2)
       tsum = Round((dl.drh.tie + dl.drb.tie) / 2)
       usum = Round((dl.drh.undone + dl.drb.undone) / 2)
       sjsum = Round((dl.drh.sj_f + dl.drh.sj_w + dl.drh.sj_t + dl.drh.sj_u + dl.drb.sj_f + dl.drb.sj_w _
              + dl.drb.sj_t + dl.drb.sj_u) / 2)
       bssum = Round((dl.drh.bs_f + dl.drh.bs_w + dl.drh.bs_t + dl.drh.bs_u + dl.drb.bs_f + dl.drb.bs_w _
               + dl.drb.bs_t + dl.drb.bs_u) / 2)
       bi = dl.mz & "共用时间：" & sjsum \ 3600 & "时" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "分" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "秒" _
            & "（占总时间" & Format(sjsum / sjz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "共用步数：" & bssum & "步" & "（占总步数" _
            & Format(bssum / bsz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "每步所用时间为" & Format(sjsum / bssum, "0.00") & "秒" _
            & "（为总用时的" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "倍）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "总共赢：输：和棋：未完成 = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "获胜比率=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "失败比率=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "和棋比率=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "未完成比率=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////
       bi = "人机模式统计：" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = dl.rj.win_
       fsum = dl.rj.fail
       tsum = dl.rj.tie
       usum = dl.rj.undone
       sjsum = dl.rj.sj_f + dl.rj.sj_w + dl.rj.sj_t + dl.rj.sj_u
       bssum = dl.rj.bs_f + dl.rj.bs_w + dl.rj.bs_t + dl.rj.bs_u
       bi = dl.mz & "共用时间：" & sjsum \ 3600 & "时" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "分" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "秒" _
            & "（占总时间" & Format(sjsum / sjz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "共用步数：" & bssum & "步" & "（占总步数" _
            & Format(bssum / bsz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "每步所用时间为" & Format(sjsum / bssum, "0.00") & "秒" _
            & "（为总用时的" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "倍）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "总共赢：输：和棋：未完成 = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "获胜比率=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "失败比率=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "和棋比率=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "未完成比率=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////
       bi = "局域网模式统计：" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = dl.wl.win_
       fsum = dl.wl.fail
       tsum = dl.wl.tie
       usum = dl.wl.undone
       sjsum = dl.wl.sj_f + dl.wl.sj_w + dl.wl.sj_t + dl.wl.sj_u
       bssum = dl.wl.bs_f + dl.wl.bs_w + dl.wl.bs_t + dl.wl.bs_u
       bi = dl.mz & "共用时间：" & sjsum \ 3600 & "时" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "分" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "秒" _
            & "（占总时间" & Format(sjsum / sjz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "共用步数：" & bssum & "步" & "（占总步数" _
            & Format(bssum / bsz, "0.00%") & "）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "每步所用时间为" & Format(sjsum / bssum, "0.00") & "秒" _
            & "（为总用时的" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "倍）" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "总共赢：输：和棋：未完成 = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "获胜比率=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "失败比率=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "和棋比率=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "未完成比率=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       Tet = sum
End If
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.index = 1 Then
   Framey.Visible = True
   Framet.Visible = False
   Command2.Visible = False
   Call xssj(1)
ElseIf TabStrip1.SelectedItem.index = 2 Then
       Framey.Visible = False
       Framet.Visible = True
       Command2.Visible = True
       Call xssj(2)
End If
End Sub

