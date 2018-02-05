VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "5"
   ClientHeight    =   1995
   ClientLeft      =   1560
   ClientTop       =   1620
   ClientWidth     =   6090
   Icon            =   "frmWINNT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6090
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdBtns 
      Caption         =   "8"
      Height          =   540
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   5895
   End
   Begin VB.CommandButton cmdBtns 
      Caption         =   "7"
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   5895
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   90
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBtns_Click(Index As Integer)
    On Error Resume Next
    UpdateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Screen Savers\Windows NT Logon", "sPicMode", Trim(Str(Index))
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then
        If FindWindow(vbNullString, LoadResString(Val(Me.Caption))) <> 0 Then End
    End If
    
    Me.Caption = LoadResString(Val(Me.Caption))
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Caption))
    Next
    
    If UCase(Left(Command, 3)) = "/C:" Then
        Dim shWnd As Long
        Dim returnV As Long
        shWnd = Val(Trim(Mid(Command, 4)))
        'MsgBox shWnd
    End If
    'A.Show vbModal, 2425986
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If UCase(Left(Command, 3)) = "/C:" Then
'        MsgBox Command
'        Dim shWnd As Long
'        Dim returnV As Long
'        shWnd = Val(Trim(Mid(Command, 4)))
'        MsgBox shWnd
'        returnV = EnableWindow(shWnd, 1)
'        MsgBox returnV
'    End If
'End Sub
