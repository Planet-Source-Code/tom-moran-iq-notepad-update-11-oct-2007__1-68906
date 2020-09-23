Attribute VB_Name = "modAPI"
Option Explicit

Public Type POINTAPI ' Declare types
    mx As Long
    my As Long
End Type

Public Pnt As POINTAPI

Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

' Win32 Declarations for Cut, Copy, Paste and Delete
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7

Public Const EM_LINEINDEX = &HBB
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_LINESCROLL = &HB6


' Win 32 Declarations for View Mode
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum

'----------------------------------------------------------------
'
'              Show fileproperties
'
'----------------------------------------------------------------
' FileName (string) is the full path and name of the file
' MyForm   (form)   is the form on wich you call the properties
' return   (long)   if return > 32 no error occured
'----------------------------------------------------------------

Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Type SHELLEXECUTEINFO
       cbSize As Long
       fMask As Long
       hwnd As Long
       lpVerb As String
       lpFile As String
       lpParameters As String
       lpDirectory As String
       nShow As Long
       hInstApp As Long
       lpIDList As Long
       lpClass As String
       hkeyClass As Long
       dwHotKey As Long
       hIcon As Long
       hProcess As Long
End Type
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long


Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
  
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsXP() As Boolean

On Error Resume Next

Dim iccex As tagInitCommonControlsEx


With iccex
  .lngSize = Len(iccex)
  .lngICC = ICC_USEREX_CLASSES
  
End With

InitCommonControlsEx iccex
InitCommonControlsXP = CBool(Err = 0)

End Function
Public Function ShowFileProp(ByVal FileName As String, aForm As Form) As Long

'if return <=32 error occured
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
If FileName = "" Then
    ShowFileProp = 0
    Exit Function
    End If
With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = aForm.hwnd
    .lpVerb = "properties"
    .lpFile = FileName
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With
r = ShellExecuteEX(SEI)
ShowFileProp = SEI.hInstApp
End Function
