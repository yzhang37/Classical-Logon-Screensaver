Attribute VB_Name = "GetDisplay"
Public Const ENUM_CURRENT_SETTINGS = (&HFFFF - 1)
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Public Type WindowsDisplay
    wdDesktopLength As Long
    wdDesktopWidth As Long
    wdDesktopSize As String
    wdColorName As String
    wdDisplayFrequency As Long
    wdColorMode As Long
    wdColorCount As Double
    wdFullDisplayInfo As String
End Type

Public Function GetDisplayMode() As WindowsDisplay
    On Error Resume Next
    Dim curDPS As DEVMODE
    Dim colors As String
    Dim i As Long
    Dim SMR As Long
    SMR = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDPS)
    
    If SMR = 0 Then
        Err.Raise 100001, , "Failed to fetch Display Card Informations."
        Exit Function
    Else
        i = curDPS.dmBitsPerPel
        GetDisplayMode.wdColorMode = i
        GetDisplayMode.wdDesktopLength = curDPS.dmPelsWidth
        GetDisplayMode.wdDesktopWidth = curDPS.dmPelsHeight
        GetDisplayMode.wdDisplayFrequency = curDPS.dmDisplayFrequency
        GetDisplayMode.wdColorCount = Int(2 ^ i)
        Select Case curDPS.dmBitsPerPel
            Case 16:     GetDisplayMode.wdColorName = "High color"
            Case 24:     GetDisplayMode.wdColorName = "True color"
            Case 32:     GetDisplayMode.wdColorName = "True color with Alpha"
            Case Else:   GetDisplayMode.wdColorName = GetDisplayMode.wdColorCount & " colors"
        End Select
        
        GetDisplayMode.wdDesktopSize = Format(curDPS.dmPelsWidth, "@@@@") & " " & Chr(&HA1C1) & " " & _
                          Format(curDPS.dmPelsHeight, "@@@@")
        GetDisplayMode.wdFullDisplayInfo = Format(curDPS.dmPelsWidth, "@@@@") & " " & Chr(&HA1C1) & " " & _
                          Format(curDPS.dmPelsHeight, "@@@@") & ", " & _
                          Format(GetDisplayMode.wdColorName, "@@@@@@@@@@@@@  ") & _
                          Format(curDPS.dmDisplayFrequency, "@@@ Hz")
    End If
End Function
