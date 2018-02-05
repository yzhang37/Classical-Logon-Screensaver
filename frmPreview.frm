VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Preview"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   89.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   109.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer mmTimer 
      Interval        =   1000
      Left            =   810
      Top             =   765
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Static iCount As Integer
    iCount = 0
    DrawIcon
End Sub

'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'    Select Case Me.BackColor
'    Case &H0
'        Me.BackColor = &H111111
'    Case &H111111
'        Me.BackColor = &H222222
'    Case &H222222
'        Me.BackColor = &H333333
'    Case &H333333
'        Me.BackColor = &H444444
'    Case &H444444
'        Me.BackColor = &H555555
'    Case &H555555
'        Me.BackColor = &H666666
'    Case &H666666
'        Me.BackColor = &H777777
'    Case &H777777
'        Me.BackColor = &H888888
'    Case &H888888
'        Me.BackColor = &H999999
'    Case &H999999
'        Me.BackColor = &HAAAAAA
'    Case &HAAAAAA
'        Me.BackColor = &HBBBBBB
'    Case &HBBBBBB
'        Me.BackColor = &HCCCCCC
'    Case &HCCCCCC
'        Me.BackColor = &HDDDDDD
'    Case &HDDDDDD
'        Me.BackColor = &HEEEEEE
'    Case &HEEEEEE
'        Me.BackColor = &HFFFFFF
'    Case &HFFFFFF
'        Me.BackColor = &H0
'    End Select
'    DrawIcon
'    ElseIf Button = 4 Then
'        MsgBox ""
'
'    End If
'End Sub

Private Sub mmTimer_Timer()
    Static iCount As Integer
    iCount = iCount + 1
    If iCount > 9 Then
        iCount = 0
        DrawIcon
    End If
End Sub

Public Sub DrawIcon()
    Cls
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
        Dim newX As Long, newY As Long
        Dim iconWidth As Integer
        Dim iconHeight As Integer
        iconWidth = 32
        iconHeight = 32
        
        x1 = Int(iconWidth / 2)
        x2 = Int(Me.ScaleWidth - x1)
        y1 = Int(iconHeight = 32 / 2)
        y2 = Int(Me.ScaleHeight - y2)
        
        If isTest = istTestFrameMode Then
            Me.ForeColor = vbWhite
            Line (x1, y1)-(x1, y2)
            Line (x1, y1)-(x2, y1)
            Line (x2, y1)-(x2, y2)
            Line (x1, y2)-(x2, y2)
        End If
            
        newX = Int(Rnd() * (x2 - x1 - 1)) + x1
        newY = Int(Rnd() * (y2 - y1 - 1)) + y1
        
        PaintIcon Me.hDc, App.EXEName & ".SCR", 0, newX, newY
End Sub
