Attribute VB_Name = "modText"
Option Explicit

 'API's and Const to load large files
 
   Private Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, lParam As Any) As Long
      
   Private Declare Function GetWindowTextLength Lib "user32" _
      Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

   Private Const WM_SETTEXT = &HC
   Private Const WM_GETTEXT = &HD
   Private Const WM_GETTEXTLENGTH = &HE
'------------------------------------------

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL As Long = &HB6

'**************************************
' For Fast Word Counting
'**************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

'Used to get Function results
Public There As Boolean
Public ret As Long


'main filename for program
Public IQFileName As String

'Globals for iQNotepad
Public PasteKey As Boolean
Public CutKey As Boolean
Public CopyKey As Boolean
Public MyMsg As String
Public CurColumn As Long
Public ToolbarOn As Boolean
Public StatusbarOn As Boolean
Public WrapOn As Boolean
Public i As Integer
Public fCancel As Boolean 'cmdlg cancel flag
Public GetForeColor As Boolean
Public TFontColor As Long
Public TBackColor As Long

Sub SaveFileAs(FileName)
    On Error Resume Next
    Dim strContents As String
    Dim FileNum As Integer
    
    Screen.MousePointer = 11
    
    FileNum = FreeFile
    
    ' Open the file.
    Open FileName For Output As #FileNum
    ' Place the contents of the notepad into a variable.
    strContents = frmMainText.Text1.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #FileNum, strContents
    Close #FileNum
    ' Reset the mouse pointer.
    Screen.MousePointer = 0

End Sub

Public Function GetColPos(tBox As Object) As Long
  GetColPos = tBox.SelStart - SendMessageByNum(tBox.hwnd, EM_LINEINDEX, -1&, 0&)
End Function
Public Function GetLineNum(tBox As Object) As Long
  GetLineNum = SendMessageByNum(tBox.hwnd, EM_LINEFROMCHAR, tBox.SelStart, 0&)
End Function
'not needed
'Public Function GetLineCount(tBox As RichTextBox) As Long
'  GetLineCount = SendMessageByNum(tBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
'End Function


Public Function LastPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  LastPart = Temp
 Else
  LastPart = Mid$(Temp, i + 1)
 End If
End Function
Public Function FirstPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  FirstPart = Temp
 Else
  FirstPart = Left$(Temp, i - 1)
 End If
End Function



Public Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function
Public Function WordCount(Text As String) As Long
    Dim dest() As Byte
    Dim i As Long


    If LenB(Text) Then
        ' Move the string's byte array into dest
        '     ()
        ReDim dest(LenB(Text))
        CopyMemory dest(0), ByVal StrPtr(Text), LenB(Text) - 1
        ' Now loop through the array and count t
        '     he words


        For i = 0 To UBound(dest) Step 2


            If dest(i) > 32 Then


                Do Until dest(i) < 33
                    i = i + 2
                Loop
                WordCount = WordCount + 1
            End If
        Next i
        Erase dest
    Else
        WordCount = 0
    End If
End Function
        

