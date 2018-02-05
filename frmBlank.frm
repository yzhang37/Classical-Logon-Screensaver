VERSION 5.00
Begin VB.Form frmBlank 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Screensaver"
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer mTime 
      Interval        =   1000
      Left            =   300
      Top             =   195
   End
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If RunMode <> rmScreenSaver Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Static iCount As Integer
    iCount = 0   '初始化计数值是否到9
    
        '如果以屏保模式运行，避免运行一个以上实例。
    If RunMode = rmScreenSaver Then
        Me.Caption = LoadResString(1) & " Screen Saver"
        If App.PrevInstance Then
        If FindWindow(vbNullString, LoadResString(Val(Me.Caption))) <> 0 Then End
        End If
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
    ShowCursor True
    End
End Sub

Private Sub mTime_Timer()
    Static iCount As Integer
    iCount = iCount + 1
    If iCount > 9 Then
        iCount = 0
        frmLogon.RandomMove
    End If
End Sub
