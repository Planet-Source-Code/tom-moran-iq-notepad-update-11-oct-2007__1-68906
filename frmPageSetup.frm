VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Setup"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Header/Footer"
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   180
      TabIndex        =   23
      Top             =   2880
      Width           =   6915
      Begin VB.TextBox txtFooter 
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Text            =   "txtFooter"
         Top             =   780
         Width           =   5535
      End
      Begin VB.TextBox txtHeader 
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Text            =   "txtHeader"
         Top             =   300
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Footer:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Header:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Printer Information"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   4320
      Width           =   6915
      Begin iQNotepad.CandyButton cmdPrintSet 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "       Set Printer..."
         Picture         =   "frmPageSetup.frx":058A
         PictureAlignment=   2
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label lblPrintername 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   3645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Printer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1740
         TabIndex        =   21
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Margins (inches)"
      ForeColor       =   &H8000000D&
      Height          =   1635
      Left            =   1920
      TabIndex        =   14
      Top             =   1080
      Width           =   3795
      Begin VB.TextBox txtBottom 
         Height          =   315
         Left            =   2580
         TabIndex        =   4
         Text            =   "txtBottom"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtTop 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Text            =   "txtTop"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtRight 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Text            =   "txtRight"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Text            =   "txtLeft"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMargins 
         AutoSize        =   -1  'True
         Caption         =   "Bottom:"
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   18
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lblMargins 
         AutoSize        =   -1  'True
         Caption         =   "Top:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   17
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label lblMargins 
         AutoSize        =   -1  'True
         Caption         =   "Right:"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   16
         Top             =   420
         Width           =   420
      End
      Begin VB.Label lblMargins 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientation"
      ForeColor       =   &H8000000D&
      Height          =   1635
      Left            =   180
      TabIndex        =   12
      Top             =   1080
      Width           =   1515
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   1275
         TabIndex        =   13
         Top             =   240
         Width           =   1275
         Begin VB.OptionButton optOrientation 
            Caption         =   "Landscape"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   660
            Width           =   1155
         End
         Begin VB.OptionButton optOrientation 
            Caption         =   "Portrait"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paper Size"
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   5595
      Begin VB.ComboBox cboPaperSize 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   360
         Width           =   4395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   345
      End
   End
   Begin iQNotepad.CandyButton cmdOkay 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Okay"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin iQNotepad.CandyButton cmdCancel 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin iQNotepad.CandyButton cmdPrint 
      Height          =   435
      Left            =   5820
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     Print"
      Picture         =   "frmPageSetup.frx":0B24
      PictureAlignment=   2
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrnError As Long


Private Sub SetPrintOptions()
On Error Resume Next
 Select Case cboPaperSize.ListIndex
 
  Case 0 'A4
   Printer.PaperSize = 9
  Case 1 'Envelope 10
   Printer.PaperSize = 20
  Case 2 'B5
   Printer.PaperSize = 34
  Case 3 'C5
   Printer.PaperSize = 28
  Case 4 'DL
   Printer.PaperSize = 27
  Case 5 'Monarch
   Printer.PaperSize = 37
  Case 6 'executive
   Printer.PaperSize = 7
  Case 7 'Legal
   Printer.PaperSize = 5
  Case 8 'Letter
   Printer.PaperSize = 1
   
 End Select
  gPaperSize = Printer.PaperSize
 
 gLeft = Val(txtLeft.Text) * 1440
 gTop = Val(txtTop.Text) * 1440
 gRight = Val(txtRight.Text) * 1440
 gBottom = Val(txtBottom.Text) * 1440
 
 gHeader = txtHeader.Text
 gFooter = txtFooter.Text
 
 If optOrientation(0).Value = True Then
  Printer.Orientation = 1
  gOrientation = 1
 Else
  Printer.Orientation = 2
  gOrientation = 2
 End If
   
   
End Sub


Private Sub cboPaperSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtLeft.SetFocus
  Exit Sub
 End If

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOkay_Click()
'set info but don't print
 
 SetPrintOptions
 
 WritePrintOptions
 
 Unload Me

End Sub

Private Sub cmdPrint_Click()
 'set info and print
 On Error Resume Next
 
 SetPrintOptions
 
 WritePrintOptions
 
  If Len(frmMainText.Text1.SelText) > 1 Then
   
   ret = MsgBox("Do you wish to print selected text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Print Text")
   
   If ret = vbCancel Then Exit Sub
   
   If ret = vbYes Then
     frmMainText.txtInsert.Text = frmMainText.Text1.SelText
     PrintRTF frmMainText.txtInsert, gLeft, gTop, gRight, gBottom   '1440 Twips = 1 Inch
     frmMainText.txtInsert.Text = ""
     Unload Me
   End If
   
   If ret = vbNo Then
     frmMainText.txtInsert.Text = frmMainText.Text1.Text
     PrintRTF frmMainText.txtInsert, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
     frmMainText.txtInsert.Text = ""
     Unload Me
   End If
  
  Else
     
     PrintRTF frmMainText.Text1, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
     Unload Me
  
  End If
 


End Sub

Private Sub cmdPrintSet_Click()
 On Error Resume Next
 frmPrnSet.Show 1
 lblPrintername.Caption = Printer.DeviceName
End Sub

Private Sub Form_Activate()
 If PrnError <> 0 Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Errhandler

 lblPrintername.Caption = Printer.DeviceName


 txtLeft.Text = Str$(gLeft / 1440)
 txtRight.Text = Str$(gRight / 1440)
 txtTop.Text = Str$(gTop / 1440)
 txtBottom.Text = Str$(gBottom / 1440)
 txtHeader.Text = gHeader
 txtFooter.Text = gFooter
 
 If gOrientation = 1 Then
  optOrientation(0).Value = True
 Else
  optOrientation(1).Value = True
 End If
 
 cboPaperSize.AddItem "A4"
 cboPaperSize.AddItem "Envelope #10"
 cboPaperSize.AddItem "Envelope B5"
 cboPaperSize.AddItem "Envelope C5"
 cboPaperSize.AddItem "Envelope DL"
 cboPaperSize.AddItem "Envelope Monarch"
 cboPaperSize.AddItem "Executive"
 cboPaperSize.AddItem "Legal"
 cboPaperSize.AddItem "Letter"
 
 Select Case gPaperSize
  
  Case 1
   cboPaperSize.Text = "Letter"
   
  Case 5
   cboPaperSize.Text = "Legal"
   
  Case 7
   cboPaperSize.Text = "Executive"
   
  Case 9
   cboPaperSize.Text = "A4"
   
  Case 20
   cboPaperSize.Text = "Envelope #10"
   
  Case 27
   cboPaperSize.Text = "Envelope DL"
      
  Case 28
   cboPaperSize.Text = "Envelope C5"
   
  Case 34
   cboPaperSize.Text = "Envelope B5"
   
  Case 37
   cboPaperSize.Text = "Envelope Monarch"
    
 End Select
Exit Sub

Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 PrnError = True
 Exit Sub
 
End Sub

Private Sub txtBottom_GotFocus()
 txtBottom.SelLength = Len(txtBottom.Text)
 
End Sub

Private Sub txtBottom_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtLeft.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtBottom_LostFocus()
If Val(txtBottom.Text) < 0.25 Then txtBottom.Text = "0.25"
End Sub

Private Sub txtFooter_Change()
 If Len(txtFooter.Text) > 60 Then
  txtFooter.Text = Left$(txtFooter.Text, 60)
  Beep
 End If
End Sub

Private Sub txtHeader_Change()
 
 If Len(txtHeader.Text) > 60 Then
  txtHeader.Text = Left$(txtHeader.Text, 60)
  Beep
 End If
 
End Sub

Private Sub txtLeft_GotFocus()
 txtLeft.SelLength = Len(txtLeft.Text)
 
End Sub


Private Sub txtLeft_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtRight.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 46 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtLeft_LostFocus()
 If Val(txtLeft.Text) < 0.25 Then txtLeft.Text = "0.25"
End Sub

Private Sub txtRight_GotFocus()
 txtRight.SelLength = Len(txtRight.Text)
 
End Sub


Private Sub txtRight_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtTop.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtRight_LostFocus()
If Val(txtRight.Text) < 0.25 Then txtRight.Text = "0.25"
End Sub

Private Sub txtTop_GotFocus()
 txtTop.SelLength = Len(txtTop.Text)
 
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then 'Enter key pressed
  KeyAscii = 0
  txtBottom.SetFocus
  Exit Sub
 End If

If KeyAscii <> 8 And KeyAscii <> 46 Then 'check for valid number or backspace key
 If Not IsNumeric(Chr(KeyAscii)) Then
  KeyAscii = 0
  MsgBox "Numbers only!", vbInformation, "Enter a number"
 End If
End If
End Sub


Private Sub txtTop_LostFocus()
If Val(txtTop.Text) < 0.25 Then txtTop.Text = "0.25"
End Sub
