Attribute VB_Name = "Module1"
Public md%, smq$(0 To 42)
Public ave%
Public Type sypw
        win_ As Integer              '赢得次数
        bs_w As Single               '赢时的步数
        sj_w As Single               '赢时的时间
        fail As Integer              '输的次数
        bs_f As Single               '失败时的步数
        sj_f As Single               '失败时的时间
        tie As Integer               '平局的次数
        bs_t As Single               '平局时的步数
        sj_t As Single               '平局时的时间
        undone As Integer            '残局次数，未完成次数
        bs_u As Single               '未完成时的步数
        sj_u As Single               '未完成时的时间
End Type
Public Type dlm         '登录信息数据类型
        mz As String * 4            '登录名
        mm As String * 10           '登录密码
        drh As sypw                 '单人时黑方输赢等等信息
        drb As sypw                 '单人时白方输赢等等信息
        rj As sypw                  '人机时输赢等等信息
        wl As sypw                  '局域网时输赢等等信息
End Type
Public Type save        '棋谱储存数据类型
        zbh As String * 4
        ysh As Single
        zbb As String * 4
        ysb As Single
        sjh As Integer
        sjb As Integer
End Type
Public dl As dlm

Sub Main()
Dim hy As dlm
Open App.Path & "zcb.lsn" For Random As #1 Len = Len(hy)
     Get #1, 1, hy
Close #1
If Trim(hy.mz) = "nocx" Then
   Formsd.Show
Else
   Formhy.Show
End If
End Sub

