Attribute VB_Name = "Reg"

' ���ģ�����ڶ���дע���ؼ��֡�
' ��ͬ��VB ���ڲ�ע�����ʷ�����������
' ͨ���ַ�����ֵ������д�κ�ע���ؼ��֡�


'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'---------------------------------------------------------------
'- ע��� Api ����...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode���ս��ַ���
Public Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Public Const REG_DWORD = 4                      ' 32-bit ����

' ע���������ֵ...
Public Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����

' ע���ؼ��ְ�ȫѡ��...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��ָ�����...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' ����ֵ...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' ���ش���
    Dim hKey As Long                                    ' ����һ��ע���ؼ���
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    
    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- ����/��ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' ����/��//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������...
    
    '------------------------------------------------------------
    '- ����/�޸Ĺؼ���ֵ...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' Ҫ��RegSetValueEx() ������Ҫ����һ���ո�...
    
    ' ����/�޸Ĺؼ���ֵ
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������
    '------------------------------------------------------------
    '- �ر�ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' �رչؼ���
    
    UpdateKey = True                                    ' ���سɹ�
    Exit Function                                       ' �˳�
CreateKeyError:
    UpdateKey = False                                   ' ���ô��󷵻ش���
    rc = RegCloseKey(hKey)                              ' ��ͼ�رչؼ���
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע���ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע���ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function



Public Function GetComputer() As String
    Dim strCompName As String
    Dim intHandle As Integer
    strCompName = String(256, " ")
    intHandle = GetComputerName(strCompName, 256)
    strCompName = Left(strCompName, InStr(1, strCompName, Chr(0)) - 1)
    GetComputer = Trim(strCompName)
End Function

