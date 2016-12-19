VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form timer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "网络时间同步~~@-@"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "timer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4710
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   2640
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   6360
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "inc@live.cn"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   2160
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   3240
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label_set 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "点我同步时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label_close 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label_net 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_net"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label_local 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_local"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1320
   End
End
Attribute VB_Name = "timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'不规则窗体
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

'移动窗体
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub Form_Load()
App.TaskVisible = True

    Dim rtn As Long
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HFF0000, 0, LWA_COLORKEY           '将扣去窗口中的蓝色
    
WebBrowser1.Navigate "http://www.timedate.cn/worldclock/ti.asp"

    Label_local.Caption = ""
    Label_net.Caption = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    'SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    '上述两种方法都能实现该功能。
End Sub

Private Sub Label_close_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_close.BackColor = vbBlack
Label_close.ForeColor = vbWhite
End Sub

Private Sub Label_close_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_close.BackColor = vbWhite
Label_close.ForeColor = vbBlack
End
End Sub

Private Sub Label_set_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_set.BackColor = vbBlack
Label_set.ForeColor = vbWhite
End Sub

Private Sub Label_set_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_set.BackColor = vbWhite
Label_set.ForeColor = vbBlack
setDateTime
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim temp As String

Label_local.Caption = "本机时间：" & Format(Now, "yyyy年mm月dd日 hh:mm:ss")

    temp = WebBrowser1.Document.body.innertext
    
    If Left(temp, 2) = "20" Then
        Label_net.Caption = "网络时间：" & Format(Now, "yyyy年mm月dd日 hh:mm:ss")
    Else
        Label_net.Caption = "网络时间获取失败" & vbCrLf & "请检查网络，然后重启本程序。"
    End If
End Sub

Sub setDateTime()
Dim temp As String
Dim netDate As String
Dim netTime As String

temp = WebBrowser1.Document.body.innertext

If Left(temp, 2) = "20" Then
    netDate = Left(temp, 4) & "-" & Mid(temp, InStr(temp, "年") + 1, InStr(temp, "月") - InStr(temp, "年") - 1) & "-" & Mid(temp, InStr(temp, "月") + 1, InStr(temp, "日") - InStr(temp, "月") - 1)
    netTime = Right(temp, 8)
    
    Cls
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         " & "提示：更新成功。"
    Print "         "
    Print "         " & "日期：" & netDate
    Print "         "
    Print "         " & "时间：" & netTime
    
    Shell "cmd.exe /c date " & netDate, vbHide
    Shell "cmd.exe /c time " & netTime, vbHide
    
Else
    Cls
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         "
    Print "         " & "提示：更新失败。"
End If
End Sub
