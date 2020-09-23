VERSION 5.00
Begin VB.Form frmMakeDefault 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make iQ Notepad Your Default Editor"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicFocus 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   225
         Picture         =   "frmMakeDefault.frx":0000
         ScaleHeight     =   1815
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   270
         Width           =   4155
         Begin VB.CheckBox ChSendto 
            Caption         =   "Place on Send to... menu"
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CheckBox ChSC 
            Caption         =   "Internet Explorer Source Code Viewer"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   960
            Width           =   3135
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Associate MS Notepad with Text (*.txt) files"
            Height          =   255
            Left            =   660
            TabIndex        =   4
            Top             =   480
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Associate iQ Notepad with Text (*.txt) files"
            Height          =   255
            Left            =   660
            TabIndex        =   3
            Top             =   120
            Width           =   3375
         End
      End
   End
   Begin iQNotepad.CandyButton cmdApply 
      Height          =   375
      Left            =   780
      TabIndex        =   7
      Top             =   2460
      Width           =   1275
      _ExtentX        =   2249
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
      Caption         =   "Apply"
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
      Left            =   2820
      TabIndex        =   8
      Top             =   2460
      Width           =   1275
      _ExtentX        =   2249
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
End
Attribute VB_Name = "frmMakeDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ASloading As Boolean 'Are we loading the form ?
Private Sub ChSC_Click()
    CheckEnabled
End Sub

Private Sub ChSendto_Click()
    CheckEnabled
End Sub

Private Sub cmdApply_Click()
    'Do the job
    If Option1.Value Then AssociateText
    If Option2.Value Then AssociateNotepad
    If ChSC.Value = 1 Then
        AddSCviewer
    Else
        RemoveSCviewer
    End If
    If ChSendto.Value = 1 Then
        AddShortCutSendTo 'create shortcut
    Else
        If FileExists(SpecialFolder(9) + "\iqnotepad.lnk") Then Kill SpecialFolder(9) + "\iqnotepad.lnk"
    End If
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me 'bail
End Sub
Private Sub Form_Load()
    ASloading = True 'loading
    Me.Icon = frmMainText.Icon
    
    'set values
    
    Option1.Value = IsAssociatedText
    Option2.Value = IsNotePadAssociatedText
    ChSC.Value = IIf(IsSCviewer, 1, 0)
    ChSendto.Value = IIf(FileExists(SpecialFolder(9) + "\iqnotepad.lnk"), 1, 0)
End Sub
Public Sub CheckEnabled()
    'compare with original states - if different enable 'Apply'
    cmdApply.Enabled = False
    If Option1.Value <> IsAssociatedText Then cmdApply.Enabled = True
    If Option2.Value <> IsNotePadAssociatedText Then cmdApply.Enabled = True
    If ChSC.Value <> IIf(IsSCviewer, 1, 0) Then cmdApply.Enabled = True
    If ChSendto.Value <> IIf(FileExists(SpecialFolder(9) + "\iqnotepad.lnk"), 1, 0) Then cmdApply.Enabled = True
End Sub
Private Sub Form_Paint()
    If ASloading Then 'OK we're loaded - do stuff we can only do once loaded
        PicFocus.SetFocus
        CheckEnabled
        ASloading = False 'Done it once - dont do it again
    End If
End Sub
Private Sub Option1_Click()
    CheckEnabled
End Sub
Private Sub Option2_Click()
    CheckEnabled
End Sub

