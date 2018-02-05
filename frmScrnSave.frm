VERSION 5.00
Begin VB.Form frmScrnSave 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   Icon            =   "frmScrnSave.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   780
   ScaleWidth      =   1200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer mTimer 
      Interval        =   1000
      Left            =   135
      Top             =   225
   End
   Begin VB.Image dispImg 
      Enabled         =   0   'False
      Height          =   225
      Left            =   495
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmScrnSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rawWid As Long
Dim rawHei As Long


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If RunMode <> rmScreenSaver Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Static iCount As Integer
    iCount = 0   '初始化计数值是否到9
    
    Dim sLOGName As String
    Dim sLOGONAME As String
    '获得颜色质量信息
    Dim disInfo As WindowsDisplay
    disInfo = GetDisplayMode()

    If disInfo.wdColorCount >= 256 Then
        sLOGName = "NTLOGONS256"
        sLOGONAME = "NTLOGO256"
    Else
        sLOGName = "NTLOGONS"
        sLOGONAME = "NTLOGO"
    End If
    
    
    '如果以屏保模式运行，避免运行一个以上实例。
    If RunMode = rmScreenSaver Then
        Me.Caption = LoadResString(1) & " Screen Saver"
        If App.PrevInstance Then
        If FindWindow(vbNullString, LoadResString(Val(Me.Caption))) <> 0 Then End
        End If
    End If

    dispImg.Stretch = False
    '先避免图片被压缩大小。


        '此处监测运行的版本
    If rMODE = RunLogonPicOnly Then
    '只有小图片
        UpdateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Screen Savers\Windows NT Logon", "sPicMode", 1
        dispImg.Picture = LoadResPicture(sLOGName, 0)
    ElseIf rMODE = RunTwoPic Then
    '有两张的图片
        Dim istr As String
        istr = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Screen Savers\Windows NT Logon", "sPicMode")

        Select Case istr
        Case "1"
            dispImg.Picture = LoadResPicture(sLOGName, 0)
        Case "0"
            dispImg.Picture = LoadResPicture(sLOGONAME, 0)
        Case Else
            UpdateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Screen Savers\Windows NT Logon", "sPicMode", 0
            dispImg.Picture = LoadResPicture(sLOGONAME, 0)
        End Select
        
        Else
        '如果是其他的版本，则不应该执行这个窗体，退出。
            End
        End If

    dispImg.Stretch = True
    rawWid = dispImg.Width
    rawHei = dispImg.Height
    '显示图片
    If RunMode = rmScreenSaver Then
        dispImg.Visible = True
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If RunMode <> rmScreenSaver Then Exit Sub
    Unload Me
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
    '运行模式为屏保模式，显示指针。
    If RunMode = rmScreenSaver Then ShowCursor 1
End Sub

Private Sub mTimer_Timer()
    Static iCount As Integer
    iCount = iCount + 1
    If iCount > 9 Then
        iCount = 0
        MovePic
    End If
End Sub

Public Sub FormPreview()
    If RunMode = rmPreview Then
        dispImg.Height = rawWid * (Me.Width / Screen.Width)
        dispImg.Width = rawHei * (Me.Height / Screen.Height)
    End If
End Sub
Public Sub MovePic()
    Cls
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim newX As Long, newY As Long
     
    x1 = Int(dispImg.Width / 2)
    x2 = Int(Me.Width - dispImg.Width / 2)
    y1 = Int(dispImg.Height / 2)
    y2 = Int(Me.Height - dispImg.Height / 2)
    
    If isTest = istTestFrameMode Then
        Me.ForeColor = vbWhite
        Line (x1, y1)-(x1, y2)
        Line (x1, y1)-(x2, y1)
        Line (x2, y1)-(x2, y2)
        Line (x1, y2)-(x2, y2)
    End If
        
    Refresh
    newX = Int(Rnd() * (x2 - x1 - 1)) + x1
    newY = Int(Rnd() * (y2 - y1 - 1)) + y1
    
    dispImg.Move Int(newX - dispImg.Width / 2), Int(newY - dispImg.Height / 2), rawWid * (Me.Width / Screen.Width), rawHei * (Me.Height / Screen.Height)
    dispImg.Visible = True
End Sub
