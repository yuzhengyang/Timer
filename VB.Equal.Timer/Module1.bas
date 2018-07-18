Attribute VB_Name = "Module1"
Public Sub helpinfo()
Dim MsgBoxInfo As Integer

    MsgBoxInfo = MsgBox("" & _
                        "作者：于正洋" & vbCrLf & vbCrLf & _
                        "邮箱：inc@live.cn" & vbCrLf & vbCrLf & _
                        "-> 开始 ： 一次开始 再次暂停" & vbCrLf & _
                        "-> 重置 ： 恢复设置的时钟" & vbCrLf & _
                        "-> 选项 ： 打开扩展选项" & vbCrLf & vbCrLf & _
                        "-> 菜单项中的 操作 建议使用 快捷键" & vbCrLf & vbCrLf & _
                        "        -> 加时 减时" & vbCrLf & _
                        "        -> 透明度 更改" & vbCrLf & _
                        "        -> 是否最前 更改" & vbCrLf & _
                        "        -> 毫秒的 隐藏 与 显示" & vbCrLf & _
                        "        -> 是否 允许超时记录" & vbCrLf & _
                        "        -> 是否 允许播放提示音乐" & vbCrLf & _
                        "        -> 同目录下音乐文件命名为 ring.mp3" & vbCrLf & vbCrLf & _
                        "改变位置 ： 拖动 窗体下半部分 空白区域" & vbCrLf & vbCrLf & _
                        "            2013 年 6 月 16 日 VER" & vbCrLf & vbCrLf & _
                        "            *-修改透明度和菜单栏", , "软件说明") '软件运行时显示帮助
End Sub

Function init(sec As Double)
    Eual.clock_state = 1
    Eual.Timer_F.Enabled = False '终止时钟
    Eual.Timer_S.Enabled = False '终止颜色变换

    Eual.Time_C.ForeColor = vbBlack  '设置毫秒标签为红色
    Eual.Time_B.ForeColor = vbBlack  '设置秒钟标签为红色
    Eual.Time_A.ForeColor = vbBlack  '设置分钟标签为红色

    'Eual.Time_A.Caption = 0 '设置分钟
    'Eual.Time_B.Caption = sec '设置秒数
    'Eual.Time_C.Caption = 0 '设置毫秒
    
    'Eual.TIMEMS = Eual.Time_A.Caption * 60 * 1000 + Eual.Time_B.Caption * 1000 + Eual.Time_C.Caption '+ 1000 '初始化时间
    Eual.TIMEMS = sec * 1000 '初始化时间
    
    Eual.Time_A.Caption = Format(Eual.TIMEMS \ 1000 \ 60, "00") '转换分钟显示
    Eual.Time_B.Caption = Format((Eual.TIMEMS - (Eual.TIMEMS \ 1000 \ 60) * 1000 * 60) \ 1000, "00") '转换秒钟显示
    Eual.Time_C.Caption = Format(Eual.TIMEMS Mod 1000, "000")  '转换毫秒显示
    
    If Dir(App.Path & "\ring.mp3") = "" Then
        '不存在的文件
    Else
        '文件存在
        'Eual.WindowsMediaPlayer1.URL = "ring.mp3"
        Eual.WindowsMediaPlayer1.Controls.stop
    End If
End Function
