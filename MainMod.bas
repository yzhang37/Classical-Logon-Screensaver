Attribute VB_Name = "MainMod"
'声明
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


'定义常数和变量
Public RunMode As Integer           '运行模式
Public rMODE As scrRUNMODEConstants '运行版本
Public Const rmConfigure = 1        '运行模式常数
Public Const rmScreenSaver = 2      '运行模式常数
Public Const rmPreview = 3          '运行模式常数

Public isTest As isTestModeConstants

Public Const HC_SYSMODALON = 4
Public Const HC_SYSMODALOFF = 5

Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GWL_HWNDPARENT = (-8)
Public Const HWND_TOP = 0
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'以下为定义的类型
Public Type OSVERSIONINFO
    dwOSVersioninfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformld As Long
    szCSDVersion As String * 128
End Type

Public Type WindowsVersion
    sLongVersion As String
    sShortVersion As String
    sCompactVersion As String
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformld As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'以下为其它子程序。

Function CompleteS(sStr As String) As String
    If Right(sStr, 1) = "\" Then
        CompleteS = sStr
        Exit Function
    End If
    CompleteS = sStr & "\"
End Function

Sub ErrorExit(Optional ByVal ownerhWnd As Long)
    MessageBoxEx ownerhWnd, LoadResString(2), LoadResString(1), vbExclamation, 0
    ShowCursor 1
    End
End Sub

Sub Main()
    On Error Resume Next
    Randomize '初始化随机种子，这是必须要做的。
    Call LoadTest
    
    Dim setHwnd As Long
    Dim runLog As String
    Dim allowConf As Boolean
    allowConf = False

    
    '获得屏幕保护的版本
    rMODE = RunLogonPicOnly
    rMODE = Val(LoadResString(102))

RESTART:
    '首先先要判断运行的版本，才能再做决定。
    Select Case rMODE
        Case 0 'RunSystem
            Dim getWin As WindowsVersion
            getWin = GetWindowsVersion()
            
            
            Select Case getWin.dwPlatformld
            Case 0                          'Windows 3.x
                rMODE = RunLogonPicOnly
                GoTo RESTART
            Case 1
                Select Case getWin.dwMinorVersion
                Case 0
                rMODE = RunLocked
                Case 10
                rMODE = RunLocked
                Case 90
                rMODE = RunLocked
                Case Else
                rMODE = RunLocked
                End Select
                GoTo RESTART
            Case 2
                Select Case getWin.dwMajorVersion
                Case 3
                rMODE = RunLogonPicOnly
                Case 4
                    rMODE = RunNT4
                Case 5
                    Select Case getWin.dwMinorVersion
                    Case 0
                        rMODE = RunLogon
                    Case 1
                        rMODE = RunLogon
                    Case 2
                        rMODE = RunLogon
                    End Select
                Case 6
                    rMODE = RunLogonPicOnly
                Case Else
                    rMODE = RunLogonPicOnly
                End Select
                GoTo RESTART
            End Select
            
        Case 1 'RunLogonWin
            
        Case 2 'RunLogonPicOnly
        Case 3 'RunTwoPic
            allowConf = True
        Case 4 'RunNT4
            runLog = """" & CompleteS(App.Path) & LoadResString(101) & """ " & Command
            GoTo RunLogon
        Case 5 'RunLogon
            runLog = """" & CompleteS(App.Path) & LoadResString(103) & """ " & Command
            GoTo RunLogon
        Case 6
            
    End Select
    
    '接着要区分运行的模式了。
    Select Case UCase(Left(Command, 2))
    Case "/S"
        'MsgBox Command, vbInformation, LoadResString(1)
        RunMode = rmScreenSaver
        If rMODE = RunLogonPicOnly Or rMODE = RunTwoPic Then
            Load frmScrnSave
            frmScrnSave.Move 0, 0, Screen.Width, Screen.Height
            frmScrnSave.MovePic
            frmScrnSave.WindowState = 2
            frmScrnSave.Show
            frmScrnSave.ZOrder 0
            ShowCursor 0     '先要隐藏鼠标指针才可以。
        ElseIf rMODE = RunLogonWin Or rMODE = RunLocked Then
            Load frmBlank
            frmBlank.Move 0, 0, Screen.Width, Screen.Height
            frmBlank.WindowState = 2
            frmBlank.Visible = True
            frmBlank.ZOrder 0
            ShowCursor 0     '先要隐藏鼠标指针才可以。
            Load frmLogon
            'frmLogon.Enabled = False
            
            '显示相应的版本内容
            If rMODE = RunLocked _
            Then
                frmLogon.SetLocked
            ElseIf _
                rMODE = RunLogonWin _
                Then
                frmLogon.SetWelcome
            Else
                ErrorExit         '出现异常错误了，退出。
            End If
            
            frmLogon.RandomMove
            frmLogon.Show
            frmLogon.ZOrder 0
        Else
            ErrorExit         '出现异常错误了，退出。
        End If
    Case "/P" 'Or ""
           Dim window_style As Long
           Dim preview_rect As RECT
           RunMode = rmPreview
           
                '得到预览区域的窗口句柄
                shWnd = Val(Trim(Mid(Command, 4)))
                'shWnd = 10093918
                    
                If rMODE = RunLogonPicOnly Or rMODE = RunTwoPic Then
                    Load frmScrnSave
                    setHwnd = frmScrnSave.hwnd
                      
                       '将窗体的Caption属性设为"Preview"
                    frmScrnSave.Caption = "Preview"
                ElseIf rMODE = RunLogonWin Or rMODE = RunLocked Then
                    Load frmPreview
                    setHwnd = frmPreview.hwnd
                      
                       '将窗体的Caption属性设为"Preview"
                    frmPreview.Caption = "Preview"
                Else
                    ErrorExit shWnd '无法继续运行。
                End If
                
                
                    '返回指定窗口客户区矩形的大小
                    GetClientRect shWnd, preview_rect
        
                    '得到frmCover窗体的信息
                    window_style = GetWindowLong(setHwnd, GWL_STYLE)
        
                    '将预览模式下的frmCover窗体style设为 子窗体 ？
                    window_style = (window_style Or WS_CHILD)
        
                    '重设frmCover窗体style为 window_style
                    SetWindowLong setHwnd, GWL_STYLE, window_style
        
                    '不懂 Set the window's parent so it appears
                    ' inside the preview area.
                    SetParent setHwnd, shWnd
        
                    '不懂 Save the preview area's hWnd in
                    ' the form's window structure.
                    SetWindowLong setHwnd, GWL_HWNDPARENT, shWnd
        
                    '不懂 Show the preview.
                    SetWindowPos setHwnd, HWND_TOP, 0&, 0&, preview_rect.Right, preview_rect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
                    
                    If rMODE = RunLogonPicOnly Or rMODE = RunTwoPic Then
                        'MessageBoxEx shWnd, "成功", "Succeed", vbInformation, 0
                        frmScrnSave.FormPreview
                        frmScrnSave.MovePic
                        frmScrnSave.MovePic
                        frmScrnSave.Refresh
                    ElseIf rMODE = RunLocked Or rMODE = RunLogonWin Then
                        frmPreview.Show
                        frmPreview.DrawIcon
                        frmPreview.Refresh
                    End If
            Exit Sub
        Case "" Or "/C"
            RunMode = rmConfigure       '运行模式为配置
            If allowConf = True Then
                frmSettings.Show        '显示配置窗口
            Else
                shWnd = Val(Trim(Mid(Command, 4)))
                MessageBoxEx shWnd, LoadResString(10), LoadResString(1), vbExclamation, 0
            End If
    End Select
    Exit Sub
RunLogon:
    Shell runLog
    Select Case Err.Number
    Case 0
        End
    Case 53
        MsgBox LoadResString(2) & vbLf & Replace(LoadResString(3), "%1", "'" & LoadResString(101) & "'"), vbExclamation, LoadResString(1)
    End Select
End Sub

Public Function GetWindowsVersion() As WindowsVersion
On Error Resume Next

Dim Osinfor As OSVERSIONINFO
Dim StrOsName As String
Osinfor.dwOSVersioninfoSize = Len(Osinfor)
GetVersionEx Osinfor

With GetWindowsVersion
    .dwBuildNumber = Osinfor.dwBuildNumber
    .dwMajorVersion = Osinfor.dwMajorVersion
    .dwMinorVersion = Osinfor.dwMinorVersion
    .dwPlatformld = Osinfor.dwPlatformld
End With

Select Case Osinfor.dwPlatformld
    Case 0
        StrOsName = "Microsoft Windows 3.2"
    Case 1
        Select Case Osinfor.dwMinorVersion
            Case 0
                StrOsName = "Microsoft Windows 95"
            Case 10
                StrOsName = "Microsoft Windows 98"
            Case 90
                StrOsName = "Microsoft Windows Millennium Edition"
        End Select
    Case 2
        Select Case Osinfor.dwMajorVersion
            Case 3
                Select Case Osinfor.dwMinorVersion
                    Case 1
                        StrOsName = "Microsoft Windows NT 3.1"
                    Case 5
                        StrOsName = "Microsoft Windows NT 3.51"
                End Select
            Case 4
                StrOsName = "Microsoft Windows NT 4.0"
            Case 5
                Select Case Osinfor.dwMinorVersion
                    Case 0
                        StrOsName = "Microsoft Windows 2000"
                    Case 1
                        StrOsName = "Microsoft Windows XP"
                    Case 2
                        StrOsName = "Microsoft Windows Server 2003"
                End Select
            Case 6
                Select Case Osinfor.dwMinorVersion
                    Case 0
                        StrOsName = "Microsoft Vista"
                    Case 1
                        StrOsName = "Microsoft Windows 7"
                    Case 2
                        StrOsName = "Microsoft Windows 8"
                    Case 3
                        StrOsName = "Microsoft Windows 8.1"
                    Case 4
                        StrOsName = "Microsoft Windows Technicial Preview"
                End Select
            Case 10
                Select Case Osinfor.dwMinorVersion
                    Case 0
                        StrOsName = "Microsoft Windows 10"
                End Select
        End Select
    Case Else
    StrOsName = "Unknow OS"
    End Select
    
    GetWindowsVersion.sLongVersion = StrOsName & ", " & LoadResString(98) & " " & Trim(Str(Osinfor.dwMajorVersion)) & "." & Trim(Str(Osinfor.dwMinorVersion)) & "." & Trim(Str(Osinfor.dwBuildNumber))
    GetWindowsVersion.sShortVersion = StrOsName
    GetWindowsVersion.sCompactVersion = Osinfor.dwMajorVersion & "." & Osinfor.dwMinorVersion & "." & Osinfor.dwBuildNumber
End Function

Private Sub LoadTest()
    On Error Resume Next
    Dim s As String
    isTest = istNormalMode
    s = LoadResString(97)
    Select Case LCase(s)
    Case "testshowframe"
        isTest = istTestFrameMode
    End Select
End Sub
