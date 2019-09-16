VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formsm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "游戏说明"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   Icon            =   "wzqcj6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9285
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog Com 
      Left            =   7320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "阅读字体设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "确 定 退 出"
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
      Height          =   2295
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Txtsm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "Formsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sm$, jj$(1 To 39)
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Com.CancelError = True
On Error GoTo errhandler
Com.Flags = cdlCFEffects Or cdlCFBoth
Com.ShowFont
Txtsm.Font.Name = Com.FontName
Txtsm.Font.Size = Com.FontSize
Txtsm.Font.Bold = Com.FontBold
Txtsm.Font.Italic = Com.FontItalic
Txtsm.Font.Underline = Com.FontUnderline
Txtsm.FontStrikethru = Com.FontStrikethru
Txtsm.ForeColor = Com.Color
Exit Sub
errhandler:
Exit Sub
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
If ave = 2 Then
   For i = 0 To 42
       sm = sm & i + 1 & ". " & smq(i) & vbCrLf & vbCrLf
   Next i
   Txtsm = sm
ElseIf ave = 1 Then
       jj(1) = "五子棋起源于古代中国，这是世界公认的。据说，五子棋最初流行于中国少数民族地区。"
       jj(2) = "五子棋起源的年代，相传是在尧舜时期，距今已有四千多年了。汉朝的班固在《弈旨》一文中阐述：“局必方正，象地则也。道必正直，神明德也。棋有白黑，阴阳分也。骈罗列布，效天文也。四象既陈，行之在人，盖王政也。”他指出了黑白棋种棋局中，棋盘线道，棋子颜色和行棋时的形象，以及行棋胜负不是靠侥幸，而是要靠人的技巧。"
       jj(3) = "五子棋将科学、艺术、竞技、娱乐与教育五者融为一体。对弈五子棋，对自己可以启智静心，与别人可以进行和谐交流。不分智商高低，老少皆宜。现今已成为深受中国乃至世界广大民众所喜爱的一项智力体育项目，已逐渐发展成为国际性文化体育交流不可缺少的内容之一。"
       jj(4) = "五子棋的下法是：对弈双方各执黑或白一色棋子，在棋盘上面，黑方先，白方后，交替落子于交叉点上。棋子落下后不能移动，没有吃子。黑方或白方哪一方先在棋盘的横、竖、斜方向的同一条直线上，由同色子连成五连为胜。"
       jj(5) = "以下为五子棋一些术语简介："
       jj(6) = "阳线：棋盘上可见的横纵直线。 阴线：棋盘上无实线连接的隐形斜线。"
       jj(7) = "活四：在棋盘某一条阳线或阴线上有同色4子不间隔地紧紧相连，且在此4子两端延长线上各有一个无子的交叉点与此4子紧密相连。"
       jj(8) = "冲四：除“活四”外的，再下一着棋便可形成五连，并且存在五连的可能性的局面。"
       jj(9) = "活三：包括“连三”和“跳三”。"
       jj(10) = "连三：在棋盘某一条阳线或阴线上有同色三子相连，且在此三子两端延长线上有一端至少有一个，另一端至少有两个无子的交叉点与此三子紧密相连。"
       jj(11) = "跳三：中间仅间隔一个无子交叉点的连三，但两端延长线均至少有一个无子的交叉点与此三子相连。"
       jj(12) = "三三：一子落下同时形成两个活三。也称“双三”。"
       jj(13) = "四四：一子落下同时形成两个冲四。也称“双四”。"
       jj(14) = "四三：一子落下同时形成一个冲四和一个活三。"
       jj(15) = "禁手：对局中禁止使用的走法。"
       jj(16) = "三三禁手：由于先走一方在无子交叉点上落子而同时形成二个或二个以上“活三”的局面，落下的这枚棋子必须是所形成的至少两个活三的共同构成子。"
       jj(17) = "四四禁手：先走一方一子落下同时形成两个（或两个以上）的冲四或活四。需要注意的是,四四禁手一个独特的地方是有在同一条线上的。"
       jj(18) = "长连禁手：先走一方落子形成长连（在阴线或者阳线上棋子连接超过五个叫长连）。"
       jj(19) = "中国现代五子棋的开拓者那威荣誉九段，多年钻研五子棋，潜心发掘五子棋的中国民间阵法，他总结了五子棋行棋的要领和临阵对局的经验，得出一套“ 秘诀 ” ，谓之《那氏五子兵法》："
       jj(20) = "        先手要攻，后手要守，以攻为守，以守待攻。"
       jj(21) = "        攻守转换，慎思变化，先行争夺，地破天惊。"
       jj(22) = "        守取外势，攻聚内力，八卦易守，成角易攻。"
       jj(23) = "        阻断分隔，稳如泰山，不思争先，胜如登天。"
       jj(24) = "        初盘争二，终局抢三，留三不冲，变化万千。"
       jj(25) = "        多个先手，细算次先，五子要点，次序在前。"
       jj(26) = "        斜线为阴，直线为阳，阴阳结合，防不胜防。"
       jj(27) = "        连三连四，易见为明，跳三跳四，暗剑深藏。"
       jj(28) = "        己落一子，敌增一兵，攻其要点，守其必争。"
       jj(29) = "        势已形成，败即降临，五子精华，一子输赢。"
       jj(30) = "职业开局名称用《彭氏口诀》进行记忆："
       jj(31) = "        二十六局先弃二，直指游星斜慧星。"
       jj(32) = "        寒星溪月疏星首，花残二月并白莲。"
       jj(33) = "        雨月金星追黑玉，松丘新宵瑞山腥。"
       jj(34) = "        星月长峡恒水流，白莲垂俏云浦岚。"
       jj(35) = "        黑玉银月倚明星，斜月名月堪称朋。"
       jj(36) = "可能很少有人注意到，五子连珠游戏其中包含着一个极为深刻的数学问题。为什么不是四子连珠，或者是六子连珠？你可能会说，四子连珠，那就太容易啦，下几步就胜了。而六子连珠呢，则太难了，谁也别想连成。这就说明，五子连珠极可能是一个最佳攻守平衡值，一个达成连珠的最大值。增一子、减一子都会打破这个平衡。四子连珠太易，攻方处于绝对优势；而六子连珠太难，守方处于绝对优势。而游戏规则必须是让游戏双方处于平等的位置才可能进行，否则游戏就不成其为游戏。要想黑白棋连珠成为一种符合游戏规则的智力游戏，五子连珠无疑是一个最佳方案。中华民族的祖先在发明五子连珠的过程中，猜想肯定也不是一蹴而就，而是极可能经历了四子连珠、六子连珠的尝试过程，最后才确定为五子连珠，并流行开来。"
       jj(37) = "为什么只给先行一方设禁手呢？也许有人认为这样的规则是不公平的，为什么只是先行方有禁手，而且有三三、四四、四三三、四四三、长连等五种之多呢? 不懂连珠的人和刚开始学连珠的人往往会认为这样的规则是不平的。然而，先落子的是先行方，这一点很关键。所以，原则上先行方可先连成五，如果紧跟着对方也连成五也不算数，还是算先行方胜，从这点看，不也是不公平吗?其实，对先行方的行棋加以限制，从对局的实际棋力的发挥来看，对双方是比较公平的。正因为先行方先下,先连五为胜，故如不对先行方加以限制，才是不公平的。请好好理解“不公平处见公平”这一奇怪的逻辑。但是，给先行方设禁手，仅仅是为了使黑白双方取得力量上的均衡吗?制定连珠规则的初期，也许目的是这样的。但随着连珠研究的发展，目前，已不仅仅是为了力量的均衡了。“四三取胜”、“不叫对方下出四三”的技术和”设置禁手”、“躲开禁手”的战术是两种截然不同的思考方法。要同时掌握这两种不同的思考方法，才更能体会出连珠的魅力。 "
       jj(38) = "！！！！！注意：五子棋实际规则为黑方先走，但此游戏可选择哪方先走，所以禁手只对先行一方起效！"
       jj(39) = "PS:此资料摘自 http://www.wuzi8.com（中国五子棋网）和百度百科的五子棋,由森哥哥收集汇总。"
       For i = 1 To 39
           sm = sm & "    " & jj(i) & vbCrLf & vbCrLf
       Next i
       Txtsm = sm
       Me.Caption = "五子棋简介"
End If
sm = ""
Txtsm.Locked = True
End Sub
