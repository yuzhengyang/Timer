VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Eual 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "ScEual"
   ClientHeight    =   6135
   ClientLeft      =   3750
   ClientTop       =   1890
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "����"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Eual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer_KeepFore 
      Interval        =   1
      Left            =   960
      Top             =   1320
   End
   Begin VB.Timer Timer_S 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer Timer_F 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1320
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2655
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8070
      _cy             =   4683
   End
   Begin VB.Label Time_A 
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Time_B 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Time_C 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   1500
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
   Begin VB.Menu start 
      Caption         =   "��ʼ"
   End
   Begin VB.Menu stop 
      Caption         =   "����"
   End
   Begin VB.Menu info 
      Caption         =   "ѡ��"
      Begin VB.Menu t_one 
         Caption         =   "�� 1 ����"
      End
      Begin VB.Menu t_three 
         Caption         =   "�� 3 ����"
      End
      Begin VB.Menu t_five 
         Caption         =   "�� 5 ����"
      End
      Begin VB.Menu e1 
         Caption         =   "==================="
         Enabled         =   0   'False
      End
      Begin VB.Menu timedown_sec 
         Caption         =   "�� ����ʱ�� - ��"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu timedown_min 
         Caption         =   "�ǩ� ����ʱ�� - ��"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu timeup_min 
         Caption         =   "�ǩ� ����ʱ�� - ��"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu timeup_sec 
         Caption         =   "�� ����ʱ�� - ��"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu e2 
         Caption         =   "==================="
         Enabled         =   0   'False
      End
      Begin VB.Menu transparentup 
         Caption         =   "͸���� - ��С"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu transparentdown 
         Caption         =   "͸���� - ����"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu e3 
         Caption         =   "==================="
         Enabled         =   0   'False
      End
      Begin VB.Menu keeppo 
         Caption         =   "�� �Ƿ� ������ǰ"
      End
      Begin VB.Menu msvisible 
         Caption         =   "�� �Ƿ� ��ʾ����"
      End
      Begin VB.Menu clock_exceed 
         Caption         =   "�� �Ƿ� ����ʱ��¼"
      End
      Begin VB.Menu musicplay 
         Caption         =   "�� �Ƿ� ������ʾ����"
      End
      Begin VB.Menu isshutdown 
         Caption         =   "�� ��ʱ ���ػ�"
      End
      Begin VB.Menu e4 
         Caption         =   "==================="
         Enabled         =   0   'False
      End
      Begin VB.Menu help 
         Caption         =   "�� ���ʹ�ð��� ��"
      End
      Begin VB.Menu yuzhengyang 
         Caption         =   "�� ������ ��"
      End
      Begin VB.Menu e5 
         Caption         =   "==================="
         Enabled         =   0   'False
      End
      Begin VB.Menu exit 
         Caption         =   "�� �˳����� ��"
      End
   End
End
Attribute VB_Name = "Eual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��������
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Private Const WS_EX_LAYERED = &H80000
'Private Const GWL_EXSTYLE = (-20)
'Private Const LWA_ALPHA = &H2
'Private Const LWA_COLORKEY = &H1

'��͸������
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'�ƶ�����
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1 '������ǰ
Const Hwndx = -1 'timerˢ����ǰ
'������ǰ

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Dim t As SYSTEMTIME '��ȡϵͳʱ��
Dim time_mark_time As Integer '��¼ʱ��
Dim time_stepcha As Integer 'ʱ���ɼ�¼ʱ�������ã�

'��ȡʱ���API

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
'���ش���߿�

Dim set_time_m As Double '����ʱ�䣨������
Dim set_transparent As Integer '����͸������

Dim clock_state_allow As Boolean '�Ƿ�����תʱ��
Dim musicprompt As Boolean '�Ƿ�����������ʾ
'**************************************************************************************
Public TIMEMS As Double 'ȫ�ֱ����ܺ�����
Public clock_state As Integer 'ȫ�ֱ���ʱ��״̬��1��������2����ʱ����״̬��

Public shutdownable As Boolean '��ʱ�����ػ�


Private Sub clock_exceed_Click()
If clock_state_allow = True Then
        clock_state_allow = False
        clock_exceed.Caption = "�� �Ƿ� ����ʱ��¼"
Else
        clock_state_allow = True
        clock_exceed.Caption = "�� �Ƿ� ����ʱ��¼"
End If

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()

'͸��͸��
set_transparent = 150
transparent (set_transparent)

'��������
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX _
, Me.Top \ Screen.TwipsPerPixelY, Me.Width \ Screen.TwipsPerPixelX, _
Me.Height \ Screen.TwipsPerPixelY, 0
    
App.TaskVisible = False

Dim H As Long
H = GetWindowLong(Me.hwnd, GWL_STYLE)
SetWindowLong Me.hwnd, GWL_STYLE, H And Not WS_CAPTION

    'Dim rtn     As Long
    'BorderStyler = 0
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, &HF0F0F0, 0, LWA_COLORKEY                '����ȥ�����е�*ɫ
    
    Me.Width = 1900
    Me.Height = 800
    
    Me.Top = Me.Width 'Screen.Height ���ó�ʼλ�á�����
    Me.Left = Screen.Width - (Me.Width * 2) '���ó�ʼλ�á�����

    Time_C.ForeColor = vbBlack  '���ú����ǩ������ɫΪ��ɫ
    Time_B.ForeColor = vbBlack  '���ú����ǩ������ɫΪ��ɫ
    Time_A.ForeColor = vbBlack  '���ú����ǩ������ɫΪ��ɫ

    Time_C.Visible = False
    
set_time_m = 300 '�����������ʱ��ģ����ӣ�

init (set_time_m)

clock_state_allow = True '����תʱ��
musicprompt = True '������ʹ��������ʾ
shutdownable = False '������ػ�


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Time_C.ForeColor = vbBlack  '���ú����ǩΪ��ɫ
'    Time_B.ForeColor = vbBlack  '�������ӱ�ǩΪ��ɫ
'    Time_A.ForeColor = vbBlack  '���÷��ӱ�ǩΪ��ɫ
'    Timer_S.Enabled = False '��ֹ��ɫ�任
    
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    'SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    '�������ַ�������ʵ�ָù��ܡ�
End Sub

Private Sub help_Click()
helpinfo
End Sub

Private Sub isshutdown_Click()
If shutdownable = False Then
    shutdownable = True
    isshutdown.Caption = "�� ��ʱ �ػ�"
Else
    shutdownable = False
    isshutdown.Caption = "�� ��ʱ ���ػ�"
End If
End Sub

Private Sub keeppo_Click()
If Timer_KeepFore.Enabled = False Then
    Timer_KeepFore.Enabled = True
    keeppo.Caption = "�� �Ƿ� ������ǰ"
Else
    Timer_KeepFore.Enabled = False
    keeppo.Caption = "�� �Ƿ� ������ǰ"
End If
End Sub

Private Sub msvisible_Click()
If Time_C.Visible = False Then
            Time_C.Visible = True
'    Time_A.Left = Me.Width / 14 * 1
'    Time_B.Left = Me.Width / 14 * 5
'    Time_C.Left = Me.Width / 14 * 9
            'Me.Height = 1111
            msvisible.Caption = "�� �Ƿ� ��ʾ����"
Else
            Time_C.Visible = False
'    Time_A.Left = Me.Width / 14 * 3
'    Time_B.Left = Me.Width / 14 * 8
            'Me.Height = 800
            msvisible.Caption = "�� �Ƿ� ��ʾ����"
End If
End Sub

Private Sub musicplay_Click()
If musicprompt = True Then
    musicplay.Caption = "�� �Ƿ� ������ʾ����"
    musicprompt = False
Else
musicplay.Caption = "�� �Ƿ� ������ʾ����"
    musicprompt = True
End If
End Sub

Private Sub start_Click()
If Timer_F.Enabled = False Then
    GetLocalTime t
    time_mark_time = t.wMilliseconds 'MARKʱ��

    Timer_F.Enabled = True
Else
    Timer_F.Enabled = False
End If
End Sub

Private Sub stop_Click()
init (set_time_m)
End Sub

Private Sub t_one_Click()
set_time_m = 60
init (set_time_m)
End Sub

Private Sub t_three_Click()
set_time_m = 180
init (set_time_m)
End Sub

Private Sub t_five_Click()
set_time_m = 300
init (set_time_m)
End Sub

Private Sub Timer_F_Timer()
GetLocalTime t
'time_mark_time = t.wMilliseconds 'MARKʱ��

If t.wMilliseconds <= time_mark_time Then
    time_stepcha = 1000 + t.wMilliseconds - time_mark_time
Else
    time_stepcha = t.wMilliseconds - time_mark_time
End If

If Eual.clock_state = 1 Then
        If TIMEMS > time_stepcha Then
            TIMEMS = TIMEMS - time_stepcha
        Else
            If clock_state_allow Then
                Eual.clock_state = 2
                
                If Dir(App.Path & "\ring.mp3") <> "" And musicprompt = True And Eual.clock_state = 2 Then
                        '�ļ�����
                        WindowsMediaPlayer1.URL = "ring.mp3"
                        'WindowsMediaPlayer1.Controls.stop
                End If
                    
                '��һ�μ�!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                '2013.6.2,�������
                TIMEMS = time_stepcha - TIMEMS
                Timer_S.Enabled = True '��ʾ������
            Else
                TIMEMS = 0
                
                    If Dir(App.Path & "\ring.mp3") <> "" And musicprompt = True Then
                        '�ļ�����
                        WindowsMediaPlayer1.URL = "ring.mp3"
                        'WindowsMediaPlayer1.Controls.stop
                    End If
                    
                Timer_F.Enabled = False '��ʱ���ر�
                Timer_S.Enabled = True '��ʾ������
            End If
        End If
Else
        If clock_state_allow Then
                TIMEMS = TIMEMS + time_stepcha
        Else
                Eual.clock_state = 1
        End If
End If


    Time_A.Caption = Format(TIMEMS \ 1000 \ 60, "00") 'ת��������ʾ
    Time_B.Caption = Format((TIMEMS - (TIMEMS \ 1000 \ 60) * 1000 * 60) \ 1000, "00") 'ת��������ʾ
    Time_C.Caption = Format(TIMEMS Mod 1000, "000")  'ת��������ʾ

time_mark_time = t.wMilliseconds 'MARKʱ��
End Sub

Private Sub Timer_KeepFore_Timer()
Dim keepfore As Long
keepfore = SetWindowPos(Me.hwnd, Hwndx, 0, 0, 0, 0, 3)
End Sub

Private Sub Timer_S_Timer()
'��ʾʱ�����
If Time_C.ForeColor = vbBlack Then     '��ʾ��ʱ......
    Time_C.ForeColor = vbRed  '���ú����ǩΪ��ɫ
    Time_B.ForeColor = vbRed  '�������ӱ�ǩΪ��ɫ
    Time_A.ForeColor = vbRed  '���÷��ӱ�ǩΪ��ɫ
Else
    Time_C.ForeColor = vbBlack  '���ú����ǩΪ��ɫ
    Time_B.ForeColor = vbBlack  '�������ӱ�ǩΪ��ɫ
    Time_A.ForeColor = vbBlack  '���÷��ӱ�ǩΪ��ɫ
End If

    If shutdownable Then
        Shell "cmd.exe /c shutdown -s -t 0"
        Timer_S.Enabled = False
    End If
End Sub



Public Sub transparent(index As Integer)
'͸��͸��
SetWindowLong hwnd, (-20), &H80000
SetLayeredWindowAttributes Me.hwnd, vbBlack, index, 2
End Sub

Private Sub timeup_min_Click()
If set_time_m < 5400 Then
set_time_m = set_time_m + 60
init (set_time_m)
End If
End Sub

Private Sub timedown_min_Click()
If set_time_m > 60 Then
set_time_m = set_time_m - 60
init (set_time_m)
End If
End Sub

Private Sub timeup_sec_Click()
If set_time_m < 5400 Then
set_time_m = set_time_m + 1
init (set_time_m)
End If
End Sub

Private Sub timedown_sec_Click()
If set_time_m > 1 Then
set_time_m = set_time_m - 1
init (set_time_m)
End If
End Sub

Private Sub transparentdown_Click()
If set_transparent > 25 Then '͸�������ӣ�����Խ��Խ͸��
set_transparent = set_transparent - 10
transparent (set_transparent)
End If
End Sub

Private Sub transparentup_Click()
If set_transparent < 245 Then '͸���ȼ��٣�Խ��Խ��͸��
set_transparent = set_transparent + 10
transparent (set_transparent)
End If
End Sub
