Attribute VB_Name = "Module1"
Public md%, smq$(0 To 42)
Public ave%
Public Type sypw
        win_ As Integer              'Ӯ�ô���
        bs_w As Single               'Ӯʱ�Ĳ���
        sj_w As Single               'Ӯʱ��ʱ��
        fail As Integer              '��Ĵ���
        bs_f As Single               'ʧ��ʱ�Ĳ���
        sj_f As Single               'ʧ��ʱ��ʱ��
        tie As Integer               'ƽ�ֵĴ���
        bs_t As Single               'ƽ��ʱ�Ĳ���
        sj_t As Single               'ƽ��ʱ��ʱ��
        undone As Integer            '�оִ�����δ��ɴ���
        bs_u As Single               'δ���ʱ�Ĳ���
        sj_u As Single               'δ���ʱ��ʱ��
End Type
Public Type dlm         '��¼��Ϣ��������
        mz As String * 4            '��¼��
        mm As String * 10           '��¼����
        drh As sypw                 '����ʱ�ڷ���Ӯ�ȵ���Ϣ
        drb As sypw                 '����ʱ�׷���Ӯ�ȵ���Ϣ
        rj As sypw                  '�˻�ʱ��Ӯ�ȵ���Ϣ
        wl As sypw                  '������ʱ��Ӯ�ȵ���Ϣ
End Type
Public Type save        '���״�����������
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

