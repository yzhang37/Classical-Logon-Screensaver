Attribute VB_Name = "ExtractIcons"
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public lIcon As Long
Public Const DI_NORMAL = &H3
Public Const DI_DEFAULTSIZE = &H8
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12

Public Function IconCount(ByVal sFileName As String) As Long
    Dim tCount As Long
    tCount = 0
    Do
        DestroyIcon lIcon
        lIcon = ExtractIcon(App.hInstance, sFileName, tCount)
        If lIcon = 0 Then Exit Do
        tCount = tCount + 1
    Loop
    IconCount = tCount
End Function

Public Sub PaintIcon(ByVal hDc As Long, ByVal sFileName As String, ByVal IconIndex As Long, ByVal x As Long, y As Long)
    DestroyIcon lIcon
    lIcon = ExtractIcon(App.hInstance, sFileName, IconIndex)
    If lIcon = 0 Then Exit Sub
    DrawIcon hDc, x, y, lIcon
End Sub

Public Sub PaintIconEx(ByVal hDc As Long, ByVal sFileName As String, ByVal IconIndex As Long, _
                        ByVal x As Long, y As Long, Optional ByVal IconSize As Long = 32, Optional ByVal DefaultSize As Boolean = True)
    DestroyIcon lIcon
    lIcon = ExtractIcon(App.hInstance, sFileName, IconIndex)
    If lIcon = 0 Then Exit Sub
    
    If DefaultSize Then
        DrawIconEx hDc, x, y, lIcon, IconSize, IconSize, 0&, 0&, DI_NORMAL Or DI_DEFAULTSIZE
    Else
        DrawIconEx hDc, x, y, lIcon, IconSize, IconSize, 0&, 0&, DI_NORMAL
    End If
End Sub

Public Function GetParentPath(opath As String) As String
    Dim t As Long
    t = Len(opath)
    
    If Right(opath, 1) = ":" Then
        If t = 2 Then
            GetParentPath = opath: Exit Function
        Else: GetParentPath = "error": Exit Function
        End If
    End If
    
    If Right(opath, 1) = "\" Then
        If t = 3 Then
            GetParentPath = opath: Exit Function
        ElseIf t < 3 Then
            GetParentPath = "error": Exit Function
        Else
            opath = Left(opath, t - 1)
        End If
    End If

    Dim tl As Long
    tl = InStrRev(opath, "\")
    GetParentPath = Left(opath, tl - 1)
End Function
