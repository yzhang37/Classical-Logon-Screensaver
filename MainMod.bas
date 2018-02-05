Attribute VB_Name = "MainMod"
'����
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


'���峣���ͱ���
Public RunMode As Integer           '����ģʽ
Public rMODE As scrRUNMODEConstants '���а汾
Public Const rmConfigure = 1        '����ģʽ����
Public Const rmScreenSaver = 2      '����ģʽ����
Public Const rmPreview = 3          '����ģʽ����

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

'����Ϊ���������
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
'����Ϊ�����ӳ���

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
    Randomize '��ʼ��������ӣ����Ǳ���Ҫ���ġ�
    Call LoadTest
    
    Dim setHwnd As Long
    Dim runLog As String
    Dim allowConf As Boolean
    allowConf = False

    
    '�����Ļ�����İ汾
    rMODE = RunLogonPicOnly
    rMODE = Val(LoadResString(102))

RESTART:
    '������Ҫ�ж����еİ汾����������������
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
    
    '����Ҫ�������е�ģʽ�ˡ�
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
            ShowCursor 0     '��Ҫ�������ָ��ſ��ԡ�
        ElseIf rMODE = RunLogonWin Or rMODE = RunLocked Then
            Load frmBlank
            frmBlank.Move 0, 0, Screen.Width, Screen.Height
            frmBlank.WindowState = 2
            frmBlank.Visible = True
            frmBlank.ZOrder 0
            ShowCursor 0     '��Ҫ�������ָ��ſ��ԡ�
            Load frmLogon
            'frmLogon.Enabled = False
            
            '��ʾ��Ӧ�İ汾����
            If rMODE = RunLocked _
            Then
                frmLogon.SetLocked
            ElseIf _
                rMODE = RunLogonWin _
                Then
                frmLogon.SetWelcome
            Else
                ErrorExit         '�����쳣�����ˣ��˳���
            End If
            
            frmLogon.RandomMove
            frmLogon.Show
            frmLogon.ZOrder 0
        Else
            ErrorExit         '�����쳣�����ˣ��˳���
        End If
    Case "/P" 'Or ""
           Dim window_style As Long
           Dim preview_rect As RECT
           RunMode = rmPreview
           
                '�õ�Ԥ������Ĵ��ھ��
                shWnd = Val(Trim(Mid(Command, 4)))
                'shWnd = 10093918
                    
                If rMODE = RunLogonPicOnly Or rMODE = RunTwoPic Then
                    Load frmScrnSave
                    setHwnd = frmScrnSave.hwnd
                      
                       '�������Caption������Ϊ"Preview"
                    frmScrnSave.Caption = "Preview"
                ElseIf rMODE = RunLogonWin Or rMODE = RunLocked Then
                    Load frmPreview
                    setHwnd = frmPreview.hwnd
                      
                       '�������Caption������Ϊ"Preview"
                    frmPreview.Caption = "Preview"
                Else
                    ErrorExit shWnd '�޷��������С�
                End If
                
                
                    '����ָ�����ڿͻ������εĴ�С
                    GetClientRect shWnd, preview_rect
        
                    '�õ�frmCover�������Ϣ
                    window_style = GetWindowLong(setHwnd, GWL_STYLE)
        
                    '��Ԥ��ģʽ�µ�frmCover����style��Ϊ �Ӵ��� ��
                    window_style = (window_style Or WS_CHILD)
        
                    '����frmCover����styleΪ window_style
                    SetWindowLong setHwnd, GWL_STYLE, window_style
        
                    '���� Set the window's parent so it appears
                    ' inside the preview area.
                    SetParent setHwnd, shWnd
        
                    '���� Save the preview area's hWnd in
                    ' the form's window structure.
                    SetWindowLong setHwnd, GWL_HWNDPARENT, shWnd
        
                    '���� Show the preview.
                    SetWindowPos setHwnd, HWND_TOP, 0&, 0&, preview_rect.Right, preview_rect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
                    
                    If rMODE = RunLogonPicOnly Or rMODE = RunTwoPic Then
                        'MessageBoxEx shWnd, "�ɹ�", "Succeed", vbInformation, 0
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
            RunMode = rmConfigure       '����ģʽΪ����
            If allowConf = True Then
                frmSettings.Show        '��ʾ���ô���
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
