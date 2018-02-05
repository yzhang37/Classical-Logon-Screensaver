VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum scrRUNMODEConstants
    RunSystem = 0
    RunLogonWin = 1
    RunLogonPicOnly = 2
    RunTwoPic = 3
    RunNT4 = 4
    RunLogon = 5
    RunLocked = 6
End Enum

Public Enum isTestModeConstants
    istNormalMode = 0
    istTestFrameMode = 1
End Enum
