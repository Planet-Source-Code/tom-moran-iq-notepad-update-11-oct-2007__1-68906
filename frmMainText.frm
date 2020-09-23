VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainText 
   AutoRedraw      =   -1  'True
   Caption         =   "iQ Notepad"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9885
   ClipControls    =   0   'False
   Icon            =   "frmMainText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9885
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "frmMainText.frx":2A52
      ScaleHeight     =   390
      ScaleWidth      =   9885
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5505
      Visible         =   0   'False
      Width           =   9885
      Begin VB.PictureBox picLnCol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6210
         Picture         =   "frmMainText.frx":1B094
         ScaleHeight     =   390
         ScaleWidth      =   4485
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   4485
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "Ln 1, Col 1"
            Height          =   195
            Left            =   2100
            TabIndex        =   24
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label lblInsNumCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INSERT"
            Height          =   195
            Index           =   0
            Left            =   140
            TabIndex        =   23
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblInsNumCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CAP"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblInsNumCap 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NUM"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1485
            TabIndex        =   21
            Top             =   120
            Width           =   375
         End
         Begin VB.Image imgGripper 
            Height          =   225
            Left            =   4220
            MousePointer    =   8  'Size NW SE
            Picture         =   "frmMainText.frx":20C3E
            Top             =   150
            Width           =   255
         End
      End
      Begin VB.Label lblIQFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Untitled"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   5700
      End
   End
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   8415
      ScaleHeight     =   1245
      ScaleWidth      =   1425
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrClipboard 
      Interval        =   50
      Left            =   9900
      Top             =   1080
   End
   Begin RichTextLib.RichTextBox txtInsert 
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   873
      _Version        =   393217
      RightMargin     =   2.00000e5
      TextRTF         =   $"frmMainText.frx":20F8C
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "frmMainText.frx":21014
      ScaleHeight     =   450
      ScaleWidth      =   9885
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   9885
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   0
         Left            =   135
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   " New "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":37856
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   1
         Left            =   720
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   " Open... "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":38268
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   2
         Left            =   1305
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   " Save "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":38802
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   3
         Left            =   1890
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   " Print "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":38D9C
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   4
         Left            =   3315
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   " Cut "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":39136
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   5
         Left            =   3900
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   " Copy "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":396D0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   6
         Left            =   4470
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   " Paste "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":39C6A
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   7
         Left            =   5040
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   " Delete "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3A204
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   8
         Left            =   6170
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " Undo "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3A79E
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   9
         Left            =   5610
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   " Find/Replace "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3AD38
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   10
         Left            =   6975
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "  Symbols "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3B2D2
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   11
         Left            =   8595
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   " Calculator "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3B86C
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   12
         Left            =   8055
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   " Change Case  "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3BE06
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   13
         Left            =   7515
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   " Spell Check "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3C3B0
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin iQNotepad.CandyButton cmdToolButton 
         Height          =   345
         Index           =   14
         Left            =   2475
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " Print Preview "
         Top             =   45
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmMainText.frx":3C94A
         Style           =   7
         Checked         =   0   'False
         ColorButtonHover=   14704640
         ColorButtonUp   =   13668448
         ColorButtonDown =   11108432
         BorderBrightness=   0
         ColorBright     =   16775930
         DisplayHand     =   0   'False
         ColorScheme     =   2
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000003&
         X1              =   6765
         X2              =   6765
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   6780
         X2              =   6780
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000003&
         X1              =   3090
         X2              =   3090
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   3105
         X2              =   3105
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         Visible         =   0   'False
         X1              =   10335
         X2              =   10335
         Y1              =   15
         Y2              =   430
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         Visible         =   0   'False
         X1              =   10320
         X2              =   10320
         Y1              =   15
         Y2              =   430
      End
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   7320
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   $"frmMainText.frx":3CEF4
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4635
      Left            =   60
      TabIndex        =   18
      Top             =   450
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8176
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmMainText.frx":3CF82
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu xxxmnuFile 
      Caption         =   "File"
      Begin VB.Menu xxxmnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu xxxmnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu xxxmnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu xxxmnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu xxmnuSepF9 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuPageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu xxxmnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu xxxmnuPrinterSetup 
         Caption         =   "Printer Setup..."
      End
      Begin VB.Menu xxxmnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu xxmnuSepF2 
         Caption         =   "-"
      End
      Begin VB.Menu xxmnuRFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuRecentFiles 
            Caption         =   " Do Not Show"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   " (Empty)"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentFiles 
            Caption         =   "RecentFiles8"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu xxmnuSepF5 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSend 
         Caption         =   "Send..."
      End
      Begin VB.Menu xxxmnuSepF7 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu xxxmnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu xxxmnuUndo 
         Caption         =   "Undo                      Ctrl+Z"
      End
      Begin VB.Menu xxmnuSepE4 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuCut 
         Caption         =   "Cut to clipboard     Ctrl+X"
      End
      Begin VB.Menu xxxmnuCopy 
         Caption         =   "Copy                      Ctrl+C"
      End
      Begin VB.Menu xxxmnuPaste 
         Caption         =   "Paste                     Ctrl+V"
      End
      Begin VB.Menu xxxmnuDelete 
         Caption         =   "Delete                    Del"
      End
      Begin VB.Menu xxmnuSepE2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuFind 
         Caption         =   "Find...                    Ctrl F"
      End
      Begin VB.Menu xxxmnuReplace 
         Caption         =   "Replace...              Ctrl+H"
      End
      Begin VB.Menu xxxmnuGoto 
         Caption         =   "Go To....                Ctrl+G"
      End
      Begin VB.Menu xxmnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSelectAll 
         Caption         =   "Select All                Ctrl+A"
      End
      Begin VB.Menu xxxmnuTimeDate 
         Caption         =   "Time/Date              F5"
      End
   End
   Begin VB.Menu xxxmnuFormat 
      Caption         =   "Format"
      Begin VB.Menu xxxmnuWordwrap 
         Caption         =   "Wordwrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu xxxmnuFont 
         Caption         =   "Font..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu xxxmnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "Font Color"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "Background Color"
      End
      Begin VB.Menu xxxmnuFSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert a text file"
      End
      Begin VB.Menu xxxmnuSaveSelected 
         Caption         =   "Save Selected Text..."
      End
      Begin VB.Menu xxxmnuFSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeCase 
         Caption         =   "Change Case"
         Begin VB.Menu mnuCaseChange 
            Caption         =   "Sentence case"
            Index           =   0
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "lower case"
            Index           =   1
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "UPPER CASE"
            Index           =   2
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "Capitalize Each Word"
            Index           =   3
         End
         Begin VB.Menu mnuCaseChange 
            Caption         =   "tOGGLE cASE"
            Index           =   4
         End
      End
   End
   Begin VB.Menu xxxmnuView 
      Caption         =   "View"
      Begin VB.Menu xxxmnuTool 
         Caption         =   "Tool Bar"
      End
      Begin VB.Menu xxxmnuStatus 
         Caption         =   "Status Bar"
      End
      Begin VB.Menu xxmnuSepV1 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuSpellCheck 
         Caption         =   "Spell Check"
         Shortcut        =   {F7}
      End
      Begin VB.Menu xxxmnuWordCount 
         Caption         =   "Document Statistics"
         Shortcut        =   ^D
      End
      Begin VB.Menu xxxmnuTextProperties 
         Caption         =   "File Properties"
      End
      Begin VB.Menu xxmnuSepV2 
         Caption         =   "-"
      End
      Begin VB.Menu xxxmnuExtended 
         Caption         =   "Extended Characters..."
         Visible         =   0   'False
      End
      Begin VB.Menu xxxmnuCharMap 
         Caption         =   "Symbols/Characters"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu xxxmnuHelpFiles 
      Caption         =   "Help"
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "iQ Notepad Manual"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "Make iQ Default Editor"
         Index           =   2
      End
      Begin VB.Menu xxxmnuHelp 
         Caption         =   "About iQ Notepad"
         Index           =   3
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmMainText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'needed for xp theme
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Findwindow is used for email sending
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

'Freeze  updates when needed
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
Private Sub OpenWordDoc(mfile As String)
    Dim WordApp As Object
    On Error GoTo Errhandler
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open mfile
    WordApp.ActiveDocument.Content.Copy
    Text1.Text = Clipboard.GetText(vbCFText)
    WordApp.Application.Quit
    Set WordApp = Nothing
    
   IQFileName = LCase(Left$(IQFileName, Len(IQFileName) - 3) & "txt")
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   lblIQFileName.Caption = LCase(IQFileName)
   Text1.DataChanged = True
   UpDateFileMenu IQFileName
   Screen.MousePointer = 0
   Exit Sub

Errhandler:
    frmConvert.Hide
    Set WordApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Private Sub InsertFile(TempFile As String)
    On Error Resume Next
    Dim FileNum As Integer
    Dim Temp As String
    
    FileNum = FreeFile

    Open TempFile For Binary As #FileNum
    
    Temp = String(LOF(FileNum), Chr$(0))
    
    Get #FileNum, , Temp
    
    Close #FileNum
    
    'check for Unicode text
    
    If Left(Temp, 2) = "ÿþ" Or Left(Temp, 2) = "þÿ" Then Temp = Replace(Right(Temp, Len(Temp) - 2), Chr(0), "")
    
    'now display
    txtInsert.Text = Temp
    Temp = ""
End Sub
Private Sub ToggleCase()

 Dim Temp As String
 Dim SChar As String
 Dim charnum As Integer
 
 'Just reverse each character in selected text
 
 For i = 1 To Len(Text1.SelText)
  SChar = Mid$(Text1.SelText, i, 1)
  charnum = Asc(SChar)
   If charnum > 96 And charnum < 123 Then
    charnum = charnum - 32
   ElseIf charnum > 64 And charnum < 91 Then
    charnum = charnum + 32
   End If
  Temp = Temp & Chr$(charnum)
 Next
 Text1.SelText = Temp
End Sub
Private Sub SentenceCase()
 Dim Temp As String
 Dim HoldText As String
 Dim SChar As String
 Dim charnum As Integer
 Dim flag As Boolean
 
 HoldText = Text1.SelText
 
 'this routine works only on lower case so be sure text is all lower
 HoldText = StrConv(HoldText, vbLowerCase)
  
 'Now parse the string to make sentence case
 For i = 1 To Len(HoldText)
  SChar = Mid$(HoldText, i, 1)
  charnum = Asc(SChar)
  
  'first letter always should be capital
  If i = 1 And charnum > 96 And charnum < 123 Then
   charnum = charnum - 32
   flag = True
  End If
 
 'Change to lower case if not first char in a sentence
 If flag = False And charnum > 64 And charnum < 91 Then
  charnum = charnum + 32
 End If
 
 'Change to uppercase if first char in a sentence
 If flag = True And charnum > 96 And charnum < 123 Then
  charnum = charnum - 32
  flag = False
 End If
 
 Temp = Temp & Chr$(charnum)
 
 'if character is period then next character will be capitalized
  If charnum = 46 Then flag = True
  
 'only for first character Flag should always be false
  If i = 1 Then flag = False
 
 
 Next i
 
 'Now assign to selected text
  Text1.SelText = Temp
 
End Sub

Private Sub FileOpen()
    On Error Resume Next
    Dim FileNum As Integer
    Dim Temp As String
    
    LockWindowUpdate Text1.hwnd
    
    FileNum = FreeFile

    Open IQFileName For Binary As #FileNum
    
    Temp = String(LOF(FileNum), Chr$(0))
    
    Get #FileNum, , Temp
    
    Close #FileNum
    
    'check for Unicode text
    
    If Left(Temp, 2) = "ÿþ" Or Left(Temp, 2) = "þÿ" Then Temp = Replace(Right(Temp, Len(Temp) - 2), Chr(0), "")
    
    'now display
    
    Text1.HideSelection = True
    DoEvents
    Text1.Text = Temp
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.SelColor = TFontColor
    DoEvents
    Text1.SelStart = 0
    Text1.SelLength = 0
    DoEvents
    Text1.HideSelection = False

    Temp = ""
    LockWindowUpdate 0
End Sub
Private Sub EditMenuEnable()
    'Enable's/disables edit menu items and toolbar buttons
    
    Dim Enabled As Boolean
    Enabled = (Text1.SelLength > 0)
    xxxmnuCut.Enabled = Enabled
    xxxmnuCopy.Enabled = Enabled
    xxxmnuDelete.Enabled = Enabled
    cmdToolButton(4).Enabled = Enabled
    cmdToolButton(5).Enabled = Enabled
    cmdToolButton(7).Enabled = Enabled
    cmdToolButton(12).Enabled = Enabled
    xxxmnuSaveSelected.Enabled = Enabled
    mnuChangeCase.Enabled = Enabled

End Sub


Private Sub PrintSelected()
 '---------------------------------------------------
 'routine to print highlighted, selected text only
 '---------------------------------------------------
 
 txtInsert.Text = Text1.SelText
 PrintRTF txtInsert, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
 txtInsert.Text = ""
 
End Sub


Private Sub RTFConvert()
On Error GoTo Errhandler

 There = FileExists(IQFileName)
 If There Then
   Screen.MousePointer = 11
   LockWindowUpdate Text1.hwnd
   txtInsert.Text = ""
   Text1.Text = ""
   txtInsert.LoadFile IQFileName, 0
   txtInsert.SelStart = 0
   txtInsert.SelLength = Len(txtInsert)
   SendMessage txtInsert.hwnd, WM_COPY, 0&, 0& 'Copy
   txtInsert.Text = ""
   Text1.SelText = Clipboard.GetText(vbCFText)
    
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.SelColor = TFontColor
    DoEvents
    Text1.SelStart = 0
    Text1.SelLength = 0
    DoEvents
    Text1.HideSelection = False
    
   IQFileName = LCase(Left$(IQFileName, Len(IQFileName) - 3) & "txt")
   
   
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   lblIQFileName.Caption = LCase(IQFileName)
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
   LockWindowUpdate 0
   Screen.MousePointer = 0
   Exit Sub
 Else
   MsgBox IQFileName & " Not Found!", vbOKOnly, "File Not Found"
   Exit Sub
 End If
 
Exit Sub

Errhandler:
 LockWindowUpdate 0
 Screen.MousePointer = 0
 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 
End Sub

Private Sub RTFNoConvert()
On Error GoTo Errhandler

 There = FileExists(IQFileName)
 If There Then
   Screen.MousePointer = 11
   LockWindowUpdate Text1.hwnd
   txtInsert.Text = ""
   Text1.Text = ""
   txtInsert.LoadFile IQFileName, 1
   txtInsert.SelStart = 0
   txtInsert.SelLength = Len(txtInsert)
   SendMessage txtInsert.hwnd, WM_COPY, 0&, 0& 'Copy
   txtInsert.Text = ""
   Text1.SelText = Clipboard.GetText(vbCFRTF)
   
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.SelColor = TFontColor
    DoEvents
    Text1.SelStart = 0
    Text1.SelLength = 0
    DoEvents
    Text1.HideSelection = False
    
   IQFileName = LCase(Left$(IQFileName, Len(IQFileName) - 3) & "txt")
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   lblIQFileName.Caption = LCase(IQFileName)
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
   LockWindowUpdate 0
   Screen.MousePointer = 0
   Exit Sub
 Else
   MsgBox IQFileName & " Not Found!", vbOKOnly, "File Not Found"
   Exit Sub
 End If
 
Exit Sub

Errhandler:
   LockWindowUpdate 0
   Screen.MousePointer = 0
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
End Sub

Private Sub UpdateLog()
 'this writes to any .LOG text files opened
 
 Text1.Text = Text1.Text & Format$(Now, "h:mm AMPM m/dd/yyyy") & vbCrLf
 'Text1.SelStart = Len(Text1.Text) + 1 '-uncomment this line to scroll to end of doc
 
End Sub


Private Sub cmdToolButton_Click(Index As Integer)

 Select Case Index
 
  Case 0 'New
   xxxmnuNew_Click
  
  Case 1 'Open
   xxxmnuOpen_Click
  
  Case 2 'Save
   xxxmnuSave_Click
  
  Case 3 'Print
   xxxmnuPrint_Click
  
  Case 4 'Cut
  
    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText
    SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
    'SendMessage Text1.hwnd, WM_CUT, 0&, 0&
  
  Case 5 'Copy
    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText
  
  Case 6 'Paste
   Text1.SelText = Clipboard.GetText
   
  Case 7 'Delete
  
    SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
    
  Case 8 'Undo
   SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
  
  Case 9 'Find/replace
   xxxmnuReplace_Click

  Case 10 'character map
   xxxmnuCharMap_Click
   
  Case 11 'calculator
   mnuCalc_Click
   
  Case 12 ' case change
   PopupMenu mnuChangeCase
   picToolbar.Refresh
   
  Case 13  'spell check
   xxxmnuSpellCheck_Click
  
  Case 14  'Print Preview
   xxxmnuPrintPreview_Click

End Select



End Sub



Private Sub Form_Activate()
 picToolbar.Refresh
End Sub

Private Sub Form_Initialize()
'  InitCommonControls
InitCommonControlsXP
End Sub

Private Sub Form_Load()
 Me.Top = (Screen.Height - Me.Height) / 2.4
 Me.Left = (Screen.Width - Me.Width) / 2
 TFontColor = QBColor(0)
 TBackColor = QBColor(15)
 
'initialise startup
picToolbar.Visible = False
ToolbarOn = False
StatusBar1.Visible = False
StatusbarOn = False
WrapOn = True
IQFileName = "Untitled"
frmMainText.Caption = " " & IQFileName & " - " & "iQ Notepad"
cmDlg.Filter = "Text (*.txt)|*.txt|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Programming (*.bas;*.frm;*.cls;*.vbp)|*.bas;*.frm;*.cls;*.vbp|Other Text (*.bat;*.csv;*.dat;*.ini;*.lst)|*.bat;*.csv;*.dat;*.ini;*.lst|Rich Text Format (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All Files (*.*)|*.*"

GetRecentFiles 'from ini file
GetOptions 'from ini file
Me.BackColor = Text1.BackColor

'printer defaults
gLeft = 1440
gRight = 1440
gTop = 1440
gBottom = 1440
gPaperSize = 1 'letter
gOrientation = 1 'portrait
gHeader = ""
gFooter = ""

GetPrintOptions 'from ini file

If WrapOn = False Then
 xxxmnuWordwrap.Checked = False
 'xxxmnuStatus.Enabled = True
 'Text1.RightMargin = 200000
 
  'If StatusbarOn = False Then
  '   StatusBar1.Visible = False
  'Else
    ' StatusBar1.Visible = True
  'End If
  
End If

 EditMenuEnable
 
'***check for auto load from command line********************************
If Len(Command$) Then
 IQFileName = Command$
 If Left$(IQFileName, 1) = Chr$(34) Then 'some programs put quotes on command line.
  IQFileName = Mid$(IQFileName, 2, Len(IQFileName) - 2)
 End If
 
 If LCase(Right$(IQFileName, 3)) = "rtf" Then
  ret = MsgBox("Do you wish to convert this file to display text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Query")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   Screen.MousePointer = 11
   Call RTFConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
  If ret = vbNo Then
   Screen.MousePointer = 11
   Call RTFNoConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
 End If
 
 There = FileExists(IQFileName)
  If There Then
   'Text1.LoadFile IQFileName, 1
   FileOpen
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
  End If
End If
'*******************************************************
lblIQFileName.Caption = LCase(IQFileName)
 
' 3rd party menu option
 'XPNetMenu1.MainBarGradientDire = usVerticality
' XPNetMenu1.CheckBack = 1743087
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Text1.DataChanged = True And Len(Text1.Text) Then
 
  ret = MsgBox("File has changed. Do you wish to save " & IQFileName & "?", vbQuestion + vbYesNoCancel, "Save?")
  
  If ret = vbYes Then
     If IQFileName = "Untitled" Then
      xxxmnuSaveAs_Click
     Else
      xxxmnuSave_Click
     End If
  ElseIf ret = vbCancel Then
     Cancel = True
     Exit Sub
  End If
  
End If
 
 Text1.Text = ""

  
 WriteOptions

End Sub

Private Sub Form_Resize()
 If Me.WindowState = 1 Then Exit Sub
 If Me.Height < 2000 Then
  Me.Height = 2000
  Me.Enabled = False
  Me.Enabled = True
 End If
 
 If Me.Width < 3000 Then
  Me.Width = 3000
  Me.Enabled = False
  Me.Enabled = True
 End If
 

 Text1.Width = Me.Width - 210
  
 If ToolbarOn = False And StatusbarOn = False Then
  Text1.Top = 0
  Text1.Height = Me.Height - 820
 End If
 
 If ToolbarOn = True And StatusbarOn = True Then
  Text1.Top = (picToolbar.Height)
  Text1.Height = Me.Height - ((picToolbar.Height + StatusBar1.Height) + 820)
  picLnCol.Left = Me.Width - 4585
 End If
  
 If ToolbarOn = True And StatusbarOn = False Then
  Text1.Top = picToolbar.Height
  Text1.Height = Me.Height - (picToolbar.Height + 820)
 End If
  
 If ToolbarOn = False And StatusbarOn = True Then
  Text1.Top = 0
  Text1.Height = Me.Height - (StatusBar1.Height + 820)
  picLnCol.Left = Me.Width - 4585
 End If
 
  picLnCol.Left = Me.Width - 4585
 
End Sub







Private Sub imgGripper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Negate VB's call to SetCapture, and tell Windows
   ' that the user is trying to resize the form.
   ReleaseCapture
   SendMessage hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub


Private Sub mnuBackColor_Click()
 GetForeColor = False
 frmColorPick.Show 1
End Sub

Private Sub mnuCalc_Click()
 
 Shell "calc.exe", vbNormalFocus
 
End Sub



Private Sub mnuCaseChange_Click(Index As Integer)
 Select Case Index
 
  Case 0 ' Sentence case
   SentenceCase
 
  Case 1 ' lower case
   Text1.SelText = StrConv(Text1.SelText, vbLowerCase)
   
  Case 2 ' upper case
   Text1.SelText = StrConv(Text1.SelText, vbUpperCase)

  Case 3 ' proper case
   Text1.SelText = StrConv(Text1.SelText, vbProperCase)
   
  Case 4 ' tOGGLE cASE
   ToggleCase

 End Select
   
     Text1.SetFocus
End Sub

Private Sub mnuFontColor_Click()
 GetForeColor = True
 frmColorPick.Show 1
End Sub

Private Sub mnuInsert_Click()
 On Error GoTo Errhandler
 Dim TempFile As String
 
 cmDlg.DialogTitle = "Insert/Merge Text File"
 cmDlg.FileName = ""
 cmDlg.Filter = "Plain Text (*.txt)|*.txt|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Other Text (*.bas;*.bat;*.csv;*.dat;*.ini;*.lst)|*.bas;*.bat;*.csv;*.dat;*.ini;*.lst|All Files (*.*)|*.*"
 cmDlg.ShowOpen
 
 Screen.MousePointer = 11
 
 TempFile = cmDlg.FileName
 
 
 There = FileExists(TempFile)
 If There Then
  txtInsert.Text = ""
  InsertFile (TempFile)
  
   txtInsert.SelStart = 0
   txtInsert.SelLength = Len(txtInsert.SelText)
   Text1.SelText = txtInsert.Text
   DoEvents
   Text1.SetFocus
 Else
   MsgBox TempFile & " Not Found!", vbOKOnly, "File Not Found"
 End If

Screen.MousePointer = 0

Exit Sub
 
Errhandler:
 Screen.MousePointer = 0
 Exit Sub
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
 
 If RecentDocs(Index) = IQFileName Then Exit Sub 'trying to open current file
 
  If Text1.DataChanged And Len(Text1.Text) Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 On Error Resume Next
 
 IQFileName = RecentDocs(Index)
 
 If LCase(Right$(IQFileName, 3)) = "rtf" Then
  ret = MsgBox("Do you wish to convert this file to display text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Query")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   Screen.MousePointer = 11
   Call RTFConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
  If ret = vbNo Then
   Screen.MousePointer = 11
   Call RTFNoConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
 End If
 
 Screen.MousePointer = 11
 There = FileExists(IQFileName)
 If There Then
   Text1.Text = ""
   FileOpen
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   lblIQFileName.Caption = LCase(IQFileName)
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
   Screen.MousePointer = 0
   Exit Sub
 Else
   Screen.MousePointer = 0
   MsgBox IQFileName & " no longer available!", vbOKOnly, "File Not Found"
   Exit Sub
 End If
End Sub

Private Sub picToolbar_Paint()
 Me.Refresh
End Sub

Private Sub Text1_Change()
 
 Text1.DataChanged = True
 picToolbar.Refresh

'only accessed if user pressed Ctrl-V paste key
If PasteKey = True Then
  SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
  SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
  Text1.SelText = Clipboard.GetText
  DoEvents
  PasteKey = False
End If

 
End Sub

Private Sub Text1_GotFocus()
 picToolbar.Refresh
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyV And Shift = 2 Then
 PasteKey = True
End If

If KeyCode = vbKeyC And Shift = 2 Then
  KeyCode = 0
  Shift = 0
  xxxmnuCopy_Click
End If

If KeyCode = vbKeyX And Shift = 2 Then
  KeyCode = 0
  Shift = 0
  xxxmnuCut_Click
End If


 If KeyCode = vbKeyH And Shift = 2 Then
  xxxmnuReplace_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

 If KeyCode = vbKeyF And Shift = 2 Then
  xxxmnuFind_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

 If KeyCode = vbKeyG And Shift = 2 Then
  xxxmnuGoto_Click
  KeyCode = 0
  Shift = 0
  Exit Sub
 End If

If KeyCode = vbKeyF5 Then
 xxxmnuTimeDate_Click
 KeyCode = 0
 Exit Sub
End If
 
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
 If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)

End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    If Button = vbRightButton Then
     PopupMenu xxxmnuEdit
     Exit Sub
    End If


End Sub


Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'for unexpected errors
On Error GoTo Errhandler


If Data.GetFormat(vbCFFiles) Then 'legit file to drop?
 'do we need to save current doc?
 If Text1.DataChanged Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 Dim OLEFileName As String
 OLEFileName = Data.Files.Item(1)
 
 If LCase(Right$(IQFileName, 3)) = "rtf" Then
  ret = MsgBox("Do you wish to convert this file to display text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Query")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   Screen.MousePointer = 11
   Call RTFConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
  If ret = vbNo Then
   Screen.MousePointer = 11
   Call RTFNoConvert
   Screen.MousePointer = 0
   Exit Sub
  End If
 End If
 
 There = FileExists(OLEFileName) 'just be sure it's there
  If There Then
   IQFileName = OLEFileName
   'Text1.LoadFile IQFileName, 1
   FileOpen
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   Text1.DataChanged = False
   If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
   UpDateFileMenu IQFileName
  Else
   MsgBox OLEFileName & " missing or invalid!", vbOKOnly + vbCritical, "iQ Notepad Drag/Drop Error"
   Exit Sub
  End If

End If

lblIQFileName.Caption = LCase(IQFileName)

Exit Sub

 

Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"

End Sub

Private Sub Text1_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub


Private Sub Text1_SelChange()

 EditMenuEnable
 
End Sub

Private Sub tmrClipboard_Timer()
    xxxmnuPaste.Enabled = Clipboard.GetFormat(vbCFText)
    cmdToolButton(6).Enabled = xxxmnuPaste.Enabled
    
    Dim b(0 To 254) As Byte

    GetKeyboardState b(0)
    If b(vbKeyNumlock) Then
       lblInsNumCap(2).Visible = True
    Else
       lblInsNumCap(2).Visible = False
    End If
    
    If b(vbKeyCapital) Then
      lblInsNumCap(1).Visible = True
    Else
       lblInsNumCap(1).Visible = False
    End If
    
    If b(vbKeyInsert) Then
      lblInsNumCap(0).Visible = False
      Else
      lblInsNumCap(0).Visible = True
    End If

End Sub



Private Sub xxxmnuCharMap_Click()
On Error GoTo Errhandler

 'display iQ charmap window
 
 frmSymbols.Show 1

 'Alternatively you could use Microsofts
 'chapmap.exe which ships with windows in sys dir.
 
 'Shell "charmap.exe", vbNormalFocus
 
 Exit Sub
 

Errhandler:

 MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Error!"
 'MsgBox "Could not find charmap.exe, character map program.", vbOKOnly, "iq Error Information"
 
 Exit Sub
 
End Sub

Private Sub xxxmnuCopy_Click()
 
  'we need to use set text so clipboard sees this as plain text not rtf text
    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText

End Sub

Private Sub xxxmnuCut_Click()

    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText
    SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
 
 
End Sub

Private Sub xxxmnuDelete_Click()

 SendMessage Text1.hwnd, WM_CLEAR, 0&, 0&
 
End Sub

Private Sub xxxmnuExit_Click()

 Unload Me
 
End Sub

Private Sub xxxmnuFind_Click()

 'Show Find API dialog
  Dim s As String

  If Text1.SelLength > 0 Then s = Text1.SelText Else s = ""
  ShowFind Me, Text1, FR_DOWN, s

  
End Sub

Private Sub xxxmnuFont_Click()
On Error GoTo Errhandler

 cmDlg.flags = cdlCFForceFontExist + cdlCFScreenFonts + cdlCFEffects
 cmDlg.Color = TFontColor
 cmDlg.FontName = Text1.Font.Name
 cmDlg.FontBold = Text1.Font.Bold
 cmDlg.FontItalic = Text1.Font.Italic
 cmDlg.FontSize = Text1.Font.Size
 cmDlg.ShowFont
 
      With cmDlg
      Text1.HideSelection = True
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1)
      
        Text1.Font.Name = .FontName
        Text1.Font.Size = .FontSize
        Text1.Font.Bold = .FontBold
        Text1.Font.Italic = .FontItalic
        Text1.SelColor = .Color
        txtInsert.Font.Name = .FontName
        txtInsert.Font.Size = .FontSize
        txtInsert.Font.Bold = .FontBold
        txtInsert.Font.Italic = .FontItalic
        'txtInsert.SelColor = .Color
        NameFont = .FontName
        SizeFont = .FontSize
        BoldFont = .FontBold
        ItalicFont = .FontItalic
        TFontColor = .Color
     Text1.SelLength = 0
     Text1.HideSelection = False
      End With
      
Errhandler:
 Exit Sub
 
End Sub

Private Sub xxxmnuGoto_Click()

 frmGoto.Show 1
 
End Sub

Private Sub xxxmnuHelp_Click(Index As Integer)

 On Error Resume Next
 Dim ftemp As String
 ftemp = App.Path & "\iqnotepad.pdf"

 Select Case Index
 
  Case 0 'iq Notepad PDF Manual if avail.
   
   There = FileExists(ftemp)
   If Not There Then
    MsgBox "iQ Notepad PDF help manual not found.  Contact iqProPlus", vbOKOnly + vbInformation, "iQ Notepad Help"
    Exit Sub
   End If
   
   ret = ShellExecute(0&, vbNullString, ftemp, vbNullString, vbNullString, vbNormalFocus)
   Exit Sub
   
  Case 2
  
   frmMakeDefault.Show 1
  
  Case 3
  
   frmAbout.Show 1
   
 End Select
  
End Sub



Private Sub xxxmnuNew_Click()
 If frmMainText.Text1.DataChanged And Len(frmMainText.Text1.Text) Then
 
  ret = MsgBox("Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "Save?")
  
  If ret = vbCancel Then Exit Sub
  
  If ret = vbNo Then
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
   Text1.DataChanged = False
   Text1.SetFocus
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    lblIQFileName.Caption = LCase(IQFileName)
   End If
   Exit Sub
  End If
  
  If IQFileName = "Untitled" Then
   xxxmnuSaveAs_Click
   If fCancel = True Then Exit Sub
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & IQFileName & " - " & "iQ Notepad"
   Text1.DataChanged = False
   Text1.SetFocus
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    lblIQFileName.Caption = LCase(IQFileName)
   End If
   Exit Sub
  Else
   xxxmnuSave_Click
   Text1.Text = ""
   IQFileName = "Untitled"
   frmMainText.Caption = " " & IQFileName & " - " & "iQ Notepad"
   Text1.DataChanged = False
   If StatusbarOn = True Then
    lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
    lblIQFileName.Caption = IQFileName
   End If
   Exit Sub
  End If
 
 Else 'it's not dirty so clear out text and reset FileName
  
  Text1.Text = ""
  Text1.DataChanged = False
  IQFileName = "Untitled"
  frmMainText.Caption = " " & IQFileName & " - " & "iQ Notepad"
  Text1.SetFocus
  If StatusbarOn = True Then
   lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
   lblIQFileName.Caption = IQFileName
  End If
  
 End If
  
End Sub


Private Sub xxxmnuOpen_Click()
 If Text1.DataChanged And Len(Text1.Text) Then
  ret = MsgBox("File has changed. Do you wish to save " & LastPart(IQFileName) & "?", vbQuestion + vbYesNoCancel, "Save?")
  If ret = vbCancel Then Exit Sub
  If ret = vbYes Then
   If IQFileName = "Untitled" Then
    xxxmnuSaveAs_Click
    If fCancel = True Then Exit Sub
   Else
    xxxmnuSave_Click
   End If
  End If
 End If
 
 On Error GoTo Errhandler
 cmDlg.Filter = "Text (*.txt)|*.txt|Html (*.htm;*.html;*.xml)|*.htm;*.html;*.xml|Log (*.log)|*.log|Programming (*.bas;*.frm;*.cls;*.vbp)|*.bas;*.frm;*.cls;*.vbp|Other Text (*.bat;*.csv;*.dat;*.ini;*.lst)|*.bat;*.csv;*.dat;*.ini;*.lst|Rich Text Format (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All Files (*.*)|*.*"
 cmDlg.ShowOpen
 
 IQFileName = cmDlg.FileName
 
 Screen.MousePointer = 11
 There = FileExists(IQFileName)
 If There Then
 'is it an RTF file?
  If LCase(Right$(IQFileName, 3)) = "rtf" Then
   ret = MsgBox("Do you wish to convert this file to display text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Query")
   If ret = vbCancel Then Exit Sub
   If ret = vbYes Then
    Screen.MousePointer = 11
    Call RTFConvert
    Screen.MousePointer = 0
    Exit Sub
   End If
   If ret = vbNo Then
    Screen.MousePointer = 11
    Call RTFNoConvert
    Screen.MousePointer = 0
    Exit Sub
   End If
  End If
 
 'is it a Word file?
  If LCase(Right$(IQFileName, 3)) = "doc" Then
     picToolbar.Refresh
     frmConvert.Show
     DoEvents
     OpenWordDoc (IQFileName)
     Unload frmConvert
     Exit Sub
  End If
 
  'None of the above so load whatever it is as plain text
    Text1.Text = ""
    FileOpen
    DoEvents
    frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
    lblIQFileName.Caption = LCase(IQFileName)
    Text1.DataChanged = False
    If Left$(Text1.Text, 4) = ".LOG" Then UpdateLog
    UpDateFileMenu IQFileName
    Screen.MousePointer = 0
    Exit Sub
 
  Else
    MsgBox IQFileName & " Not Found!", vbOKOnly, "File Not Found"
    Screen.MousePointer = 0
    Exit Sub
  End If
 
Errhandler:
  Screen.MousePointer = 0
  Exit Sub
 
End Sub

Private Sub xxxmnuPageSetup_Click()
 On Error GoTo Errhandler
 Dim prncheck As Variant
 
 prncheck = Printer.DeviceName

 frmPageSetup.Show 1
 
 Exit Sub
 
Errhandler:
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub
   
End Sub

Private Sub xxxmnuPaste_Click()
 Text1.SelText = Clipboard.GetText
End Sub

Private Sub xxxmnuPrint_Click()
Dim prncheck As Variant

' print
 On Error GoTo Errhandler

 prncheck = Printer.DeviceName

  If Len(Text1.SelText) > 1 Then
   
   ret = MsgBox("Do you wish to print selected text only?", vbYesNoCancel + vbQuestion, "iQ Notepad Print Text")
   
   If ret = vbCancel Then Exit Sub
   
   If ret = vbYes Then
    Call PrintSelected
    Exit Sub
   End If
   
   If ret = vbNo Then
    txtInsert.Text = Text1.Text
    PrintRTF txtInsert, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
    frmMainText.txtInsert.Text = ""
   End If
   
  Else
    txtInsert.Text = Text1.Text
    PrintRTF txtInsert, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
    frmMainText.txtInsert.Text = ""
    
  End If

  
  
 Exit Sub

Errhandler:
     
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub
   
End Sub

Private Sub xxxmnuPrinterSetup_Click()
 

'call print cmdlg
 On Error GoTo Errhandler
 
 cmDlg.flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDUseDevModeCopies
 cmDlg.ShowPrinter

 Printer.Copies = cmDlg.Copies
 
 For i = 1 To Printer.Copies
  Printer.Orientation = gOrientation
  Printer.PaperSize = gPaperSize
  PrintRTF Text1, gLeft, gTop, gRight, gBottom  '1440 Twips = 1 Inch
 Next
 
 Exit Sub
 
Errhandler:
     
     If Err <> 32755 Then '32755 is cancel error
     
       MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
       
     End If
  Screen.MousePointer = 0
End Sub


Private Sub xxxmnuPrintPreview_Click()
 On Error GoTo Errhandler
 Dim prncheck As Variant
 
 prncheck = Printer.DeviceName
 
    gPrint = False
    
    frmPrintPreview.Show 1
    
    If gPrint Then
     xxxmnuPrint_Click
     gPrint = False
    End If
 
 Exit Sub
 
Errhandler:
   MsgBox Err.Number & " " & Error$, vbOKOnly + vbCritical, "iQ Print Error!"
   Exit Sub

End Sub

Private Sub xxxmnuReplace_Click()

 'Show Find/Replace API dialog
  Dim s As String
 
   If Text1.SelLength > 0 Then s = Text1.SelText Else s = ""
   ShowFind Me, Text1, 0, s, True, ""

  
End Sub

Private Sub xxxmnuSave_Click()
 On Error GoTo Errhandler

 
 If IQFileName = "Untitled" Then
  xxxmnuSaveAs_Click
  Exit Sub
 End If

 SaveFileAs IQFileName
 Text1.DataChanged = False
 Exit Sub
 
Errhandler:

 MsgBox "Error" & Str$(Err) & " - " & Error$, vbOKOnly + vbCritical, "File not saved!"
 Exit Sub
 
End Sub


Private Sub xxxmnuSaveAs_Click()
 On Error GoTo Errhandler
 fCancel = False
 
 cmDlg.FileName = IQFileName
 cmDlg.flags = cdlOFNOverwritePrompt
 
 cmDlg.ShowSave
 
 IQFileName = cmDlg.FileName
 
 SaveFileAs IQFileName
 Text1.DataChanged = False
 
 frmMainText.Caption = " " & LastPart(IQFileName) & " - " & "iQ Notepad"
 lblIQFileName.Caption = LCase(IQFileName)
 UpDateFileMenu IQFileName
 
 Exit Sub
 
Errhandler:
 If Err = 32755 Then fCancel = True
 If Err <> 32755 Then
  MsgBox "Error" & Str$(Err) & " - " & Error$, vbOKOnly + vbCritical, "File not saved!"
 End If

 Exit Sub

End Sub

Private Sub xxxmnuSaveSelected_Click()
 On Error GoTo Errhandler
 fCancel = False
 Dim TempFile As String
 
 cmDlg.flags = cdlOFNOverwritePrompt
 
 cmDlg.ShowSave
 
 TempFile = cmDlg.FileName

 Open TempFile For Output As 1
  Print #1, Text1.SelText
 Close 1

 
 Exit Sub
 
Errhandler:
 If Err = 32755 Then fCancel = True
 If Err <> 32755 Then
  MsgBox "Error" & Str$(Err) & " - " & Error$, vbOKOnly + vbCritical, "File not saved!"
 End If

 Exit Sub
End Sub

Private Sub xxxmnuSelectAll_Click()
 
 Text1.SelStart = 0
 Text1.SelLength = Len(Text1)
 

End Sub

Private Sub xxxmnuSend_Click()
 Dim Start As Long
 If IQFileName = "Untitled" Then
  xxxmnuSaveAs_Click
  If fCancel = True Then Exit Sub
 Else
  xxxmnuSave_Click
 End If
 

   ' Start Outlook Express email.
   
    ShellExecute Me.hwnd, "Open", _
        "mailto:?subject=" & IQFileName & "&body=File Attached", _
        vbNullString, vbNullString, vbNormalFocus

    ' Wait until Outlook Express is ready.
    While ret = 0
        DoEvents
        ret = FindWindow(vbNullString, IQFileName)
    Wend
    
    Start = Timer + 0.3
    While Start > Timer
     DoEvents
    Wend
    ' Send keys Alt-I-A, the zip file name,
    ' two TABs, and Enter.
    SendKeys "%ia" & IQFileName & "{TAB}{TAB}{ENTER}"
    
    
End Sub

Private Sub xxxmnuSpellCheck_Click()
'--------------------------------
'Code to use WSpell 3rd party control
'--------------------------------

'  On Error GoTo OpenError
 ' Dim Scount As Long
  
  'Check to see if checking string or whole doc
'  If Text1.SelLength > 1 Then
    'checking selected text only
    'Dim text As String

    ' Get the contents of the text box into a string.
    
    'text = Text1.SelText
 '   txtInsert.Text = Text1.SelText
 '   WSpell1.TextControlHWnd = txtInsert.hwnd

    

    ' Check the spelling of the string. Note that WSpell1's ShowContext
    ' property is set to True. This causes
    ' a context display to appear in the spell-check dialog box.
    'WSpell1.ShowContext = True
    'WSpell1.text = text
 '   ret = WSpell1.Start
 '   If (ret >= 0) Then
        ' The user didn't cancel, so put the correct string back into the text box.
 '       Text1.SelText = txtInsert.Text
 '   End If
 '  Else
    'check whole document
 '   WSpell1.ShowContext = False
 '   WSpell1.TextControlHWnd = Text1.hwnd
 '   Call WSpell1.Start
 '  End If
 '   Scount = WSpell1.WordsReplacedCount
 '   MyMsg = "Spell check completed." & vbCrLf
 '   MyMsg = MyMsg & "Words replaced:" & Str$(Scount) & ".  "
    
 '   MsgBox MyMsg, vbOKOnly + vbInformation, "iQ WordPad"
 '   Exit Sub


'Code for using MS Word Spell Check
    
    Dim oWord As Object
    Dim oTmpDoc As Object
    Dim lOrigTop As Long
'
    On Error GoTo OpenError
    Screen.MousePointer = 11
    
    ' Create a Word document object...
    Set oWord = CreateObject("Word.Application")
    Set oTmpDoc = oWord.Documents.Add
    oWord.Visible = False
    
   ' Position Word off screen to avoid having document visible...
    oWord.WindowState = 0
    oWord.Top = -3000
    LockWindowUpdate Text1.hwnd
    Text1.HideSelection = True
    
    ' copy the contents of the text box to the clipboard
    If Text1.SelLength < 2 Then
     Text1.SelStart = 0
     Text1.SelLength = Len(Text1.Text)
    End If
    
    'be sure to use Text so no formatting
    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText

    ' Assign the text to the document and check spelling...
    Screen.MousePointer = 0
    With oTmpDoc
        .Content.Paste
        .Activate
        .CheckSpelling
       
       ' .CheckGrammar 'uncomment this line to include grammarcheck
       
       ' After user has made changes, use the clipboard to
       ' transfer the contents back to the text box
        .Content.Copy
        
         Text1.SelText = Clipboard.GetText(vbCFText)

        ' Close the document and exit Word...
       .Saved = True
        Clipboard.Clear
       .Close
      End With
      
      'release the object
       Set oTmpDoc = Nothing
       oWord.Quit
       Set oWord = Nothing
      
      'Now tell the user we're done
       LockWindowUpdate 0
       SendKeys "{BACKSPACE}" 'get rid of extra return/linefeed added by word
       DoEvents
       MsgBox "Spell check is complete.", vbOKOnly + vbInformation, "iQ WordPad"
       Text1.HideSelection = False

      Exit Sub
      
'==========================================
'code for Spell Check Anywhere program
'==========================================
' If Len(Text1.TextRTF) < 2 Then Exit Sub
' Clipboard.Clear
 
' If Text1.SelLength < 2 Then
'   Text1.SelStart = 3
'   Text1.SetFocus
'   DoEvents
' End If
 
' SendKeys "{F11}", True
'
 'For i = 1 To 10000
' Next
 
' DoEvents
' Exit Sub

OpenError:
 MsgBox Error$ & " - " & Err, vbOKOnly, "Spell Check Error"
 Text1.HideSelection = False
 LockWindowUpdate 0
 Exit Sub
 
End Sub

Private Sub xxxmnuStatus_Click()
 If StatusbarOn = False Then
  StatusbarOn = True
  xxxmnuStatus.Checked = True
  StatusBar1.Visible = True
  Form_Resize
  If StatusbarOn = True Then lblStatus.Caption = "Ln " & Format(GetLineNum(frmMainText.Text1) + 1, "###,###,###,###") & ", Col " & (GetColPos(frmMainText.Text1) + 1)
 Else
  StatusbarOn = False
  xxxmnuStatus.Checked = False
  StatusBar1.Visible = False
  Form_Resize
 End If
End Sub

Private Sub xxxmnuTextProperties_Click()
 'show property dialog
 ret = ShowFileProp(IQFileName, frmMainText)
 
End Sub

Private Sub xxxmnuTimeDate_Click()
 'Text1.SelText = Format$(Now, "h:mm AMPM m/dd/yyyy")
 frmDateTime.Show 1
End Sub

Private Sub xxxmnuTool_Click()
 If ToolbarOn = False Then
  ToolbarOn = True
  xxxmnuTool.Checked = True
  picToolbar.Visible = True
  Form_Resize
  picToolbar.Refresh
 Else
  ToolbarOn = False
  xxxmnuTool.Checked = False
  picToolbar.Visible = False
  Form_Resize
 End If
End Sub

Private Sub xxxmnuUndo_Click()

 SendMessage Text1.hwnd, EM_UNDO, 0&, 0&
 
End Sub

Private Sub xxxmnuWordCount_Click()
On Error Resume Next
 Dim WCount As Long
 Dim LnCount As Long
 Dim CharCount As Long
 CharCount = 0
 WCount = WordCount(Text1.Text)
 LnCount = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0, 0&)
 CharCount = SendMessageLong(Text1.hwnd, WM_GETTEXTLENGTH, 0, 0)
 CharCount = Format(CharCount, "###,###,###,###,###")
 MyMsg = ""
 MyMsg = IQFileName & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Word Count:" & Str$(WCount) & "   " & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Line Count:" & Str$(LnCount) & "   " & vbCrLf & vbCrLf
 MyMsg = MyMsg & "   Character Count:" & Str$(CharCount) & "   " & vbCrLf & vbCrLf
 
 MsgBox MyMsg, vbOKOnly + vbInformation, "iQ Notepad Statistics"

End Sub

Private Sub xxxmnuWordwrap_Click()
    xxxmnuWordwrap.Checked = Not xxxmnuWordwrap.Checked
   ' DoEvents
    WrapOn = xxxmnuWordwrap.Checked
    
   Text1.RightMargin = IIf(xxxmnuWordwrap.Checked, 0, 200000)
    
    
   'xxxmnuStatus.Enabled = IIf(xxxmnuWordwrap.Checked, 0, -1)
    
   If xxxmnuStatus.Checked = True And xxxmnuStatus.Enabled Then
     StatusBar1.Visible = True
    StatusbarOn = True
   Else
     StatusBar1.Visible = False
     StatusbarOn = False
    End If

    Form_Resize
End Sub


