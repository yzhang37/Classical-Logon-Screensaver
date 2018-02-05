VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log on Windows"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10080
   ControlBox      =   0   'False
   Icon            =   "fProvIcon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ctlPics 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1380
      Index           =   1
      Left            =   4470
      ScaleHeight     =   1320
      ScaleWidth      =   5460
      TabIndex        =   3
      Tag             =   "14"
      Top             =   105
      Visible         =   0   'False
      Width           =   5520
      Begin VB.Label ctlLbls 
         AutoSize        =   -1  'True
         Caption         =   "Tips"
         Height          =   180
         Index           =   3
         Left            =   975
         TabIndex        =   5
         Tag             =   "15"
         Top             =   135
         Width           =   360
      End
      Begin VB.Label ctlLbls 
         AutoSize        =   -1  'True
         Caption         =   "Press C-A-D Message"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Tag             =   "18"
         Top             =   780
         Width           =   1710
      End
   End
   Begin VB.PictureBox ctlPics 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1530
      Index           =   0
      Left            =   315
      ScaleHeight     =   1470
      ScaleWidth      =   5460
      TabIndex        =   0
      Tag             =   "11"
      Top             =   1260
      Visible         =   0   'False
      Width           =   5520
      Begin VB.Label ctlLbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1650
         TabIndex        =   6
         Tag             =   "104"
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label ctlLbls 
         AutoSize        =   -1  'True
         Caption         =   "Why Press C-A-D"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Tag             =   "12"
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label ctlLbls 
         AutoSize        =   -1  'True
         Caption         =   "Press C-A-D Message"
         Height          =   180
         Index           =   0
         Left            =   975
         TabIndex        =   1
         Tag             =   "13"
         Top             =   135
         Width           =   1710
      End
   End
   Begin VB.Image imgNTLogo 
      Enabled         =   0   'False
      Height          =   810
      Index           =   0
      Left            =   30
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const UnitV = 60

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If RunMode <> rmScreenSaver Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim gWin As WindowsVersion
    Dim i As Long
    Dim sFontName As String
    Dim compName As String
    Dim userNam As String
    Dim disInfo As WindowsDisplay
    disInfo = GetDisplayMode()
    
    Dim sLOGName As String
    Dim sSlp As String
    
    If disInfo.wdColorCount >= 256 Then
        sLOGName = "NTLOGONS256"
        sSlp = "NTSLP256"
    Else
        sLOGName = "NTLOGONS"
        sSlp = "NTSLIP16"
    End If
    imgNTLogo(0).Move 0, 0
    imgNTLogo(0).Picture = LoadResPicture(sLOGName, 0)
    Me.Width = imgNTLogo(0).Width + Me.Width - Me.ScaleWidth
    Load imgNTLogo(1)
    imgNTLogo(1).Picture = LoadResPicture(sSlp, 0)
    imgNTLogo(1).Stretch = False
    imgNTLogo(1).Move imgNTLogo(0).Left, imgNTLogo(0).Height
    imgNTLogo(1).Visible = True

    gWin = GetWindowsVersion()
    sFontName = LoadResString(107)
    If gWin.dwPlatformld = 2 Then
        If gWin.dwMajorVersion >= 5 Then sFontName = LoadResString(105)
    End If
    
    For Each Control In Me.Controls
        Control.FontName = sFontName
        Control.FontSize = Val(LoadResString(106))
        Control.Caption = LoadResString(Val(Control.Tag))
        Control.Tag = LoadResString(Val(Control.Tag))
    Next
    
    compName = GetComputer()
    userNam = VBA.Environ$("username")
    
    If Len(userNam) > 0 Then
    ctlLbls(2).Caption = Replace(ctlLbls(2).Caption, "%1", compName)
    ctlLbls(2).Caption = Replace(ctlLbls(2).Caption, "%2", userNam)
    Else
    ctlLbls(2).Caption = Replace(ctlLbls(2).Caption, "%1\%2", LoadResString(19))
    End If

    Select Case gWin.dwPlatformld
    Case 0              'Windows 3.x
        GoTo WIN3
    Case 1              'Windows 9x
        Select Case gWin.dwMinorVersion
            Case 0
                GoTo 95
            Case 10
                GoTo 98
            Case 90
                GoTo 98
            Case Else
                GoTo 95
            End Select
    Case 2
    '处理 NT 系列的 Windows 图标打印问题。
    
            Select Case gWin.dwMajorVersion
            Case 3
                GoTo WIN3
            Case 4
                GoTo 98
            Case 5
                GoTo NT5
            Case 6
                GoTo NT6
            Case Else
                GoTo 95
            End Select
            
    End Select
    GoTo Cont
WIN3:
    PaintIcon ctlPics(0).hDc, App.EXEName & ".SCR", 3, 10, 5
    GoTo Cont
95:
    PaintIcon ctlPics(0).hDc, App.EXEName & ".SCR", 2, 10, 5
    GoTo Cont
98:
    PaintIcon ctlPics(0).hDc, "main.cpl", 6, 10, 5
    GoTo Cont
NT5:
    PaintIcon ctlPics(0).hDc, "main.cpl", 7, 10, 5
    GoTo Cont
NT6:
    PaintIcon ctlPics(0).hDc, "main.cpl", 5, 10, 5
    GoTo Cont
Cont:
    PaintIcon ctlPics(1).hDc, App.EXEName & ".SCR", 1, 10, 5
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    End
End Sub

Public Sub SetLocked()
    With imgNTLogo(1)
        ctlPics(1).Move .Left, .Top + .Height, .Width
    End With
    Me.Caption = ctlPics(1).Tag
    ctlLbls(2).Left = ctlLbls(3).Left

    
    '设置自动换行，以免超屏幕
    ctlLbls(3).Width = ctlPics(1).ScaleWidth - ctlLbls(3).Left
    ctlLbls(3).WordWrap = True
    
    ctlLbls(2).Width = ctlPics(1).ScaleWidth - ctlLbls(2).Left
    ctlLbls(2).WordWrap = True
    '====
    
    ctlLbls(2).Top = ctlLbls(3).Top + ctlLbls(3).Height * 3 + UnitV * 3
    
    ctlPics(1).BorderStyle = 0
    ctlPics(1).Visible = True
    ctlPics(1).Height = ctlLbls(2).Top + ctlLbls(2).Height + 2.5 * UnitV
    Me.Height = Me.Height - Me.ScaleHeight + ctlPics(1).Top + ctlPics(1).Height + 2 * UnitV
End Sub

Public Sub SetWelcome()
    With imgNTLogo(1)
        ctlPics(0).Move .Left, .Top + .Height, .Width
    End With
    Me.Caption = ctlPics(0).Tag
    ctlLbls(1).Left = ctlLbls(0).Left

    
    '设置自动换行，以免超屏幕
    ctlLbls(0).Width = ctlPics(0).ScaleWidth - ctlLbls(1).Left
    ctlLbls(0).WordWrap = True
    
    ctlLbls(1).Width = ctlPics(0).ScaleWidth - ctlLbls(1).Left
    ctlLbls(1).WordWrap = True
    '====
    ctlLbls(1).Top = ctlLbls(0).Top + ctlLbls(0).Height * 3 + UnitV * 3
    
    
    
    ctlLbls(4).Top = ctlLbls(1).Top + ctlLbls(1).Height + UnitV
    ctlLbls(4).Left = ctlPics(0).Width - ctlLbls(4).Width - UnitV
    ctlLbls(4).ForeColor = vbBlue
    ctlPics(0).BorderStyle = 0

    ctlPics(0).Height = ctlLbls(4).Top + ctlLbls(4).Height + UnitV
    ctlPics(0).Visible = True
    Me.Height = Me.Height - Me.ScaleHeight + ctlPics(0).Top + ctlPics(0).Height + 2 * UnitV
End Sub

Public Sub RandomMove()
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim newX As Long
    Dim newY As Long
    x1 = Int(Me.Width / 2)
    y1 = Int(Me.Height / 2)
    x2 = Screen.Width - x1
    y2 = Screen.Height - y1
    
    If isTest = istTestFrameMode Then
        frmBlank.Line (x1, y1)-(x1, y2)
        frmBlank.Line (x1, y1)-(x2, y1)
        frmBlank.Line (x2, y1)-(x2, y2)
        frmBlank.Line (x1, y2)-(x2, y2)
    End If
    newX = Int(Rnd() * (x2 - x1 - 1)) + x1
    newY = Int(Rnd() * (y2 - y1 - 1)) + y1
    
    Me.Move Int(newX - Me.Width / 2), Int(newY - Me.Height / 2)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static x0 As Integer
    Static y0 As Integer
    If RunMode <> rmScreenSaver Then Exit Sub
    
    If ((x0 = 0) And (y0 = 0)) Or _
        ((Abs(x0 - x) < 5) And (Abs(y0 - y) < 5)) _
        Then
            x0 = x
            y0 = y
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor True
    End
End Sub

