Attribute VB_Name = "Module1"
Public Sub helpinfo()
Dim MsgBoxInfo As Integer

    MsgBoxInfo = MsgBox("" & _
                        "���ߣ�������" & vbCrLf & vbCrLf & _
                        "���䣺inc@live.cn" & vbCrLf & vbCrLf & _
                        "-> ��ʼ �� һ�ο�ʼ �ٴ���ͣ" & vbCrLf & _
                        "-> ���� �� �ָ����õ�ʱ��" & vbCrLf & _
                        "-> ѡ�� �� ����չѡ��" & vbCrLf & vbCrLf & _
                        "-> �˵����е� ���� ����ʹ�� ��ݼ�" & vbCrLf & vbCrLf & _
                        "        -> ��ʱ ��ʱ" & vbCrLf & _
                        "        -> ͸���� ����" & vbCrLf & _
                        "        -> �Ƿ���ǰ ����" & vbCrLf & _
                        "        -> ����� ���� �� ��ʾ" & vbCrLf & _
                        "        -> �Ƿ� ����ʱ��¼" & vbCrLf & _
                        "        -> �Ƿ� ��������ʾ����" & vbCrLf & _
                        "        -> ͬĿ¼�������ļ�����Ϊ ring.mp3" & vbCrLf & vbCrLf & _
                        "�ı�λ�� �� �϶� �����°벿�� �հ�����" & vbCrLf & vbCrLf & _
                        "            2013 �� 6 �� 16 �� VER" & vbCrLf & vbCrLf & _
                        "            *-�޸�͸���ȺͲ˵���", , "���˵��") '�������ʱ��ʾ����
End Sub

Function init(sec As Double)
    Eual.clock_state = 1
    Eual.Timer_F.Enabled = False '��ֹʱ��
    Eual.Timer_S.Enabled = False '��ֹ��ɫ�任

    Eual.Time_C.ForeColor = vbBlack  '���ú����ǩΪ��ɫ
    Eual.Time_B.ForeColor = vbBlack  '�������ӱ�ǩΪ��ɫ
    Eual.Time_A.ForeColor = vbBlack  '���÷��ӱ�ǩΪ��ɫ

    'Eual.Time_A.Caption = 0 '���÷���
    'Eual.Time_B.Caption = sec '��������
    'Eual.Time_C.Caption = 0 '���ú���
    
    'Eual.TIMEMS = Eual.Time_A.Caption * 60 * 1000 + Eual.Time_B.Caption * 1000 + Eual.Time_C.Caption '+ 1000 '��ʼ��ʱ��
    Eual.TIMEMS = sec * 1000 '��ʼ��ʱ��
    
    Eual.Time_A.Caption = Format(Eual.TIMEMS \ 1000 \ 60, "00") 'ת��������ʾ
    Eual.Time_B.Caption = Format((Eual.TIMEMS - (Eual.TIMEMS \ 1000 \ 60) * 1000 * 60) \ 1000, "00") 'ת��������ʾ
    Eual.Time_C.Caption = Format(Eual.TIMEMS Mod 1000, "000")  'ת��������ʾ
    
    If Dir(App.Path & "\ring.mp3") = "" Then
        '�����ڵ��ļ�
    Else
        '�ļ�����
        'Eual.WindowsMediaPlayer1.URL = "ring.mp3"
        Eual.WindowsMediaPlayer1.Controls.stop
    End If
End Function
