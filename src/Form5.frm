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
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog Com 
      Left            =   3360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�� �� �� �� ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ȷ �� �� ��"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "������ģʽ"
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
         Caption         =   "�˻�ģʽ"
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
         Caption         =   "����ģʽ"
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
            Caption         =   "�׷���"
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
            Caption         =   "�ڷ���"
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
            Caption         =   " ԭ ʼ �� ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " ͳ �� �� ��"
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
Formsj.Caption = dl.mz & "����Ϸ��ս����"
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
Labdhy.Caption = "�ܹ���Ӯ" & dl.drh.win_ & "��" & vbCrLf & "����" & _
dl.drh.bs_w & "��" & vbCrLf & "����" & dl.drh.sj_w \ 3600 & "ʱ" _
& (dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600) \ 60 & "��" & _
dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600 - ((dl.drh.sj_w - (dl.drh.sj_w \ 3600) * 3600) \ 60) * 60 & "��"
Labdhs.Caption = "�ܹ�����" & dl.drh.fail & "��" & vbCrLf & "����" & _
dl.drh.bs_f & "��" & vbCrLf & "����" & dl.drh.sj_f \ 3600 & "ʱ" _
& (dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600) \ 60 & "��" & _
dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600 - ((dl.drh.sj_f - (dl.drh.sj_f \ 3600) * 3600) \ 60) * 60 & "��"
Labdhh.Caption = "�ܹ��Ѻ���" & dl.drh.tie & "��" & vbCrLf & "����" & _
dl.drh.bs_t & "��" & vbCrLf & "����" & dl.drh.sj_t \ 3600 & "ʱ" _
& (dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600) \ 60 & "��" & _
dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600 - ((dl.drh.sj_t - (dl.drh.sj_t \ 3600) * 3600) \ 60) * 60 & "��"
Labdhw.Caption = "�ܹ�δ������" & dl.drh.undone & "��" & vbCrLf & "����" & _
dl.drh.bs_u & "��" & vbCrLf & "����" & dl.drh.sj_u \ 3600 & "ʱ" _
& (dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600) \ 60 & "��" & _
dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600 - ((dl.drh.sj_u - (dl.drh.sj_u \ 3600) * 3600) \ 60) * 60 & "��"
'///////////////////////////////////////////////////////////////////////////
Labdby.Caption = "�ܹ���Ӯ" & dl.drb.win_ & "��" & vbCrLf & "����" & _
dl.drb.bs_w & "��" & vbCrLf & "����" & dl.drb.sj_w \ 3600 & "ʱ" _
& (dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600) \ 60 & "��" & _
dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600 - ((dl.drb.sj_w - (dl.drb.sj_w \ 3600) * 3600) \ 60) * 60 & "��"
Labdbs.Caption = "�ܹ�����" & dl.drb.fail & "��" & vbCrLf & "����" & _
dl.drb.bs_f & "��" & vbCrLf & "����" & dl.drb.sj_f \ 3600 & "ʱ" _
& (dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600) \ 60 & "��" & _
dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600 - ((dl.drb.sj_f - (dl.drb.sj_f \ 3600) * 3600) \ 60) * 60 & "��"
Labdbh.Caption = "�ܹ��Ѻ���" & dl.drb.tie & "��" & vbCrLf & "����" & _
dl.drb.bs_t & "��" & vbCrLf & "����" & dl.drb.sj_t \ 3600 & "ʱ" _
& (dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600) \ 60 & "��" & _
dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600 - ((dl.drb.sj_t - (dl.drb.sj_t \ 3600) * 3600) \ 60) * 60 & "��"
Labdbw.Caption = "�ܹ�δ������" & dl.drb.undone & "��" & vbCrLf & "����" & _
dl.drb.bs_u & "��" & vbCrLf & "����" & dl.drb.sj_u \ 3600 & "ʱ" _
& (dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600) \ 60 & "��" & _
dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600 - ((dl.drb.sj_u - (dl.drb.sj_u \ 3600) * 3600) \ 60) * 60 & "��"
'///////////////////////////////////////////////////////////////////////////
Labr.Caption = dl.mz & "��"
Labry.Caption = "�ܹ���Ӯ" & dl.rj.win_ & "��" & vbCrLf & "����" & _
dl.rj.bs_w & "��" & vbCrLf & "����" & dl.rj.sj_w \ 3600 & "ʱ" _
& (dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600) \ 60 & "��" & _
dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600 - ((dl.rj.sj_w - (dl.rj.sj_w \ 3600) * 3600) \ 60) * 60 & "��"
Labrs.Caption = "�ܹ�����" & dl.rj.fail & "��" & vbCrLf & "����" & _
dl.rj.bs_f & "��" & vbCrLf & "����" & dl.rj.sj_f \ 3600 & "ʱ" _
& (dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600) \ 60 & "��" & _
dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600 - ((dl.rj.sj_f - (dl.rj.sj_f \ 3600) * 3600) \ 60) * 60 & "��"
Labrh.Caption = "�ܹ��Ѻ���" & dl.rj.tie & "��" & vbCrLf & "����" & _
dl.rj.bs_t & "��" & vbCrLf & "����" & dl.rj.sj_t \ 3600 & "ʱ" _
& (dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600) \ 60 & "��" & _
dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600 - ((dl.rj.sj_t - (dl.rj.sj_t \ 3600) * 3600) \ 60) * 60 & "��"
Labrw.Caption = "�ܹ�δ������" & dl.rj.undone & "��" & vbCrLf & "����" & _
dl.rj.bs_u & "��" & vbCrLf & "����" & dl.rj.sj_u \ 3600 & "ʱ" _
& (dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600) \ 60 & "��" & _
dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600 - ((dl.rj.sj_u - (dl.rj.sj_u \ 3600) * 3600) \ 60) * 60 & "��"
'/////////////////////////////////////////////////////////////////////////
Labw.Caption = dl.mz & "��"
Labwy.Caption = "�ܹ���Ӯ" & dl.wl.win_ & "��" & vbCrLf & "����" & _
dl.wl.bs_w & "��" & vbCrLf & "����" & dl.wl.sj_w \ 3600 & "ʱ" _
& (dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600) \ 60 & "��" & _
dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600 - ((dl.wl.sj_w - (dl.wl.sj_w \ 3600) * 3600) \ 60) * 60 & "��"
Labws.Caption = "�ܹ�����" & dl.wl.fail & "��" & vbCrLf & "����" & _
dl.wl.bs_f & "��" & vbCrLf & "����" & dl.wl.sj_f \ 3600 & "ʱ" _
& (dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600) \ 60 & "��" & _
dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600 - ((dl.wl.sj_f - (dl.wl.sj_f \ 3600) * 3600) \ 60) * 60 & "��"
Labwh.Caption = "�ܹ��Ѻ���" & dl.wl.tie & "��" & vbCrLf & "����" & _
dl.wl.bs_t & "��" & vbCrLf & "����" & dl.wl.sj_t \ 3600 & "ʱ" _
& (dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600) \ 60 & "��" & _
dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600 - ((dl.wl.sj_t - (dl.wl.sj_t \ 3600) * 3600) \ 60) * 60 & "��"
Labww.Caption = "�ܹ�δ������" & dl.wl.undone & "��" & vbCrLf & "����" & _
dl.wl.bs_u & "��" & vbCrLf & "����" & dl.wl.sj_u \ 3600 & "ʱ" _
& (dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600) \ 60 & "��" & _
dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600 - ((dl.wl.sj_u - (dl.wl.sj_u \ 3600) * 3600) \ 60) * 60 & "��"
ElseIf i = 2 Then
       On Error Resume Next
       Dim wsum%, fsum%, tsum%, usum%, sjsum!, sjz!, bssum!, bsz!, bi$, dh$
       dh = "������������������������������������������������������������" & vbCrLf
       bi = "����ͳ�ƣ�" & vbCrLf
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
       bi = dl.mz & "����ʱ�䣺" & sjsum \ 3600 & "ʱ" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "��" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "���ò�����" & bssum & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ÿ������ʱ��Ϊ" & Format(sjsum / bssum, "0.00") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�ܹ�Ӯ���䣺���壺δ��� = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "��ʤ����=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ʧ�ܱ���=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�������=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "δ��ɱ���=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////////
       bi = "����ģʽͳ�ƣ�" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = Round((dl.drh.win_ + dl.drb.win_) / 2)
       fsum = Round((dl.drh.fail + dl.drb.fail) / 2)
       tsum = Round((dl.drh.tie + dl.drb.tie) / 2)
       usum = Round((dl.drh.undone + dl.drb.undone) / 2)
       sjsum = Round((dl.drh.sj_f + dl.drh.sj_w + dl.drh.sj_t + dl.drh.sj_u + dl.drb.sj_f + dl.drb.sj_w _
              + dl.drb.sj_t + dl.drb.sj_u) / 2)
       bssum = Round((dl.drh.bs_f + dl.drh.bs_w + dl.drh.bs_t + dl.drh.bs_u + dl.drb.bs_f + dl.drb.bs_w _
               + dl.drb.bs_t + dl.drb.bs_u) / 2)
       bi = dl.mz & "����ʱ�䣺" & sjsum \ 3600 & "ʱ" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "��" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "��" _
            & "��ռ��ʱ��" & Format(sjsum / sjz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "���ò�����" & bssum & "��" & "��ռ�ܲ���" _
            & Format(bssum / bsz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ÿ������ʱ��Ϊ" & Format(sjsum / bssum, "0.00") & "��" _
            & "��Ϊ����ʱ��" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "����" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�ܹ�Ӯ���䣺���壺δ��� = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "��ʤ����=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ʧ�ܱ���=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�������=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "δ��ɱ���=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////
       bi = "�˻�ģʽͳ�ƣ�" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = dl.rj.win_
       fsum = dl.rj.fail
       tsum = dl.rj.tie
       usum = dl.rj.undone
       sjsum = dl.rj.sj_f + dl.rj.sj_w + dl.rj.sj_t + dl.rj.sj_u
       bssum = dl.rj.bs_f + dl.rj.bs_w + dl.rj.bs_t + dl.rj.bs_u
       bi = dl.mz & "����ʱ�䣺" & sjsum \ 3600 & "ʱ" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "��" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "��" _
            & "��ռ��ʱ��" & Format(sjsum / sjz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "���ò�����" & bssum & "��" & "��ռ�ܲ���" _
            & Format(bssum / bsz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ÿ������ʱ��Ϊ" & Format(sjsum / bssum, "0.00") & "��" _
            & "��Ϊ����ʱ��" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "����" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�ܹ�Ӯ���䣺���壺δ��� = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "��ʤ����=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ʧ�ܱ���=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�������=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "δ��ɱ���=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       '////////////////////////////////////////////////////////////////////////////////
       bi = "������ģʽͳ�ƣ�" & vbCrLf
       sum = sum & dh & bi & dh
       wsum = dl.wl.win_
       fsum = dl.wl.fail
       tsum = dl.wl.tie
       usum = dl.wl.undone
       sjsum = dl.wl.sj_f + dl.wl.sj_w + dl.wl.sj_t + dl.wl.sj_u
       bssum = dl.wl.bs_f + dl.wl.bs_w + dl.wl.bs_t + dl.wl.bs_u
       bi = dl.mz & "����ʱ�䣺" & sjsum \ 3600 & "ʱ" _
            & (sjsum - (sjsum \ 3600) * 3600) \ 60 & "��" & _
            sjsum - (sjsum \ 3600) * 3600 - ((sjsum - (sjsum \ 3600) * 3600) \ 60) * 60 & "��" _
            & "��ռ��ʱ��" & Format(sjsum / sjz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "���ò�����" & bssum & "��" & "��ռ�ܲ���" _
            & Format(bssum / bsz, "0.00%") & "��" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ÿ������ʱ��Ϊ" & Format(sjsum / bssum, "0.00") & "��" _
            & "��Ϊ����ʱ��" & Format((sjsum / bssum) / (sjz / bsz), "0.00") & "����" & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�ܹ�Ӯ���䣺���壺δ��� = " & wsum & ":" & fsum & ":" & tsum & ":" & usum & vbCrLf
       sum = sum & bi
       bi = dl.mz & "��ʤ����=" & Format(wsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "ʧ�ܱ���=" & Format(fsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "�������=" & Format(tsum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
       sum = sum & bi
       bi = dl.mz & "δ��ɱ���=" & Format(usum / (wsum + fsum + tsum + usum), "0.00%") & vbCrLf
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

