Attribute VB_Name = "ModAPI"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const WM_GETTEXT = &HD         'Text of List Box Item
Public Const WM_GETTEXTLENGTH = &HE   'Length of List Box Item
Public Const WM_KEYDOWN = &H100       'Key is Down
Public Const WM_KEYUP = &H101         'Key is Up

Public Const LB_GETCOUNT = &H18B      'Number of List Box Items
Public Const LB_GETTEXT = &H189       'Text of List Box Item
Public Const LB_SETCURSEL = &H186     'Set Current Item of List Box

Public Const VK_UP = &H26             'Up Arrow is Press
Public Const VK_DOWN = &H28           'Down Arrow is Press
