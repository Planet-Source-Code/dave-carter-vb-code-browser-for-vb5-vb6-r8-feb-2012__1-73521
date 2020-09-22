VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   10
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00FFFFFF&
      Height          =   3555
      HelpContextID   =   14
      Index           =   3
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   6915
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CheckBox chkCodeAutoEncode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto RTF Encoding in Code Viewer (Syntax Colouring)"
         Height          =   375
         Left            =   3660
         TabIndex        =   55
         Top             =   2700
         Value           =   1  'Checked
         Width           =   3195
      End
      Begin VB.CheckBox chkCodeLoadAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load entire Code File from Project Explorer Selection"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   2700
         Width           =   3195
      End
      Begin VB.CheckBox chkTipOfTheDay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tip of the Day on Start Up"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   2220
         Width           =   3195
      End
      Begin VB.CheckBox chkShowCHeadCount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Project / Class Explorer Head Count"
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   3795
      End
      Begin VB.CheckBox chkConfirmExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Confirm Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   4995
      End
      Begin VB.CheckBox chkHideTool 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Toolbar on Project Child Forms"
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   900
         Width           =   4995
      End
      Begin VB.CheckBox chkHideProj 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide Project Explorer on Project Child Forms"
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   480
         Width           =   4995
      End
      Begin VB.CheckBox chkUseChildWindows 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Child Windows"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   60
         Width           =   2535
      End
      Begin VB.Label lblLink 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MySpace"
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   4
         Left            =   6060
         TabIndex        =   56
         Tag             =   "http://www.myspace.com/dave_e_c"
         Top             =   3180
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00FFFFFF&
      Height          =   3555
      HelpContextID   =   12
      Index           =   1
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   6915
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CheckBox chkAttributes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attributes"
         Height          =   375
         Left            =   90
         TabIndex        =   32
         ToolTipText     =   "Show Attributes"
         Top             =   2640
         Width           =   1515
      End
      Begin VB.ComboBox cboFont 
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   420
         Width           =   1695
      End
      Begin VB.ComboBox cboFont 
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   585
      End
      Begin VB.CheckBox chkBold 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bold"
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   900
         Width           =   675
      End
      Begin VB.CheckBox chkLineNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Line Numbers"
         Height          =   375
         Left            =   90
         TabIndex        =   33
         ToolTipText     =   "Show Line NUmbers"
         Top             =   2940
         Width           =   1515
      End
      Begin VB.CheckBox chkSyntaxColours 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use own colours"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2895
         Left            =   2700
         TabIndex        =   34
         Top             =   420
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   5106
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         TextRTF         =   $"frmOptions.frx":058A
      End
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Attributes"
         Height          =   255
         Index           =   3
         Left            =   1860
         TabIndex        =   28
         Top             =   1260
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Numbers"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1260
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Text"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblSynColours 
         BackStyle       =   0  'Transparent
         Caption         =   "Viewer Back Colour"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1920
         TabIndex        =   21
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1920
         TabIndex        =   23
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1920
         TabIndex        =   25
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1920
         TabIndex        =   27
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1920
         TabIndex        =   29
         Top             =   2700
         Width           =   645
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1920
         TabIndex        =   31
         Top             =   3000
         Width           =   645
      End
   End
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00FFFFFF&
      Height          =   3555
      HelpContextID   =   13
      Index           =   2
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   6915
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton cmdUnzipFolder 
         Caption         =   "..."
         Height          =   375
         Left            =   6240
         TabIndex        =   43
         ToolTipText     =   "Nominate Unzip Folder"
         Top             =   2520
         Width           =   495
      End
      Begin VB.CheckBox chkAutoCleanUnzipFolder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto-Clean Unzip Folder"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1500
         Width           =   3435
      End
      Begin VB.CheckBox chkAutoLoadVBP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Load VBP from Unzip"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   780
         Width           =   2415
      End
      Begin VB.CheckBox chkAutoUnzip 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto-UnZip"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label lblLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Unzip Help on the web"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   47
         Tag             =   "http://www.opensource.apple.com/source/zip/zip-6/unzip/unzip/windll/windll.txt?txt"
         ToolTipText     =   "Click to view documentation about using Unzip32, on the web"
         Top             =   3195
         Width           =   1815
      End
      Begin VB.Label lblUnzipCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unzip Folder (Nominal until used)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   2340
      End
      Begin VB.Label lblUnzipInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Check this to automatically load VBP/s found in Zip File following an UnZip."
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   1140
         Width           =   5955
      End
      Begin VB.Label lblLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Info-Zip Home Page"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   46
         Tag             =   "http://www.info-zip.org/"
         ToolTipText     =   "Click to visit Info-Zip Home Page"
         Top             =   2895
         Width           =   1815
      End
      Begin VB.Label lblLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Unzip dll @ vbAccelerator"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   45
         Tag             =   "http://www.vbaccelerator.com/home/VB/Code/Libraries/Compression/Unzipping_Files/article.asp"
         ToolTipText     =   "Click to Download Unzip32 from vbAccelerator"
         Top             =   3195
         Width           =   1995
      End
      Begin VB.Label lblUnzipFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "Unzip Folder"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Unzip Folder is nominal until used"
         Top             =   2580
         Width           =   6015
      End
      Begin VB.Label lblUnzipInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Check this to automatically Unzip Zip Files to Unzip Folder"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   420
         Width           =   5955
      End
      Begin VB.Label lblUnzipInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Check this to automatically delete items in the Unzip Folder on Program Exit."
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   5955
      End
      Begin VB.Label lblLink 
         BackStyle       =   0  'Transparent
         Caption         =   "Unzip dll @ CodeGuru"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   44
         Tag             =   "http://www.codeguru.com/vb/gen/vb_graphics/fileformats/article.php/c6743"
         ToolTipText     =   "Click to Download Unzip32 from CodeGuru"
         Top             =   2895
         Width           =   1815
      End
   End
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00FFFFFF&
      Height          =   3555
      HelpContextID   =   11
      Index           =   0
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   6915
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   6975
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Index           =   3
         Left            =   6300
         TabIndex        =   13
         Top             =   2940
         Width           =   495
      End
      Begin VB.CheckBox chkShowMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Open in Text Editor"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2580
         Width           =   2415
      End
      Begin VB.CheckBox chkShowMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Open with VB4"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   2415
      End
      Begin VB.CheckBox chkShowMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Open with VB5"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chkShowMenu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Open with VB6"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   6300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Index           =   1
         Left            =   6300
         TabIndex        =   7
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Index           =   2
         Left            =   6300
         TabIndex        =   10
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Editor Path"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   5955
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "VB6 Path"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2220
         Width           =   5955
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "VB5 Path"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   5955
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "VB4 Path"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   5955
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4395
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7752
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Project Menu"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Syntax Colours"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "UnZip"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   59
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6300
      TabIndex        =   60
      Top             =   4800
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdgColour 
      Left            =   840
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Item Colour"
   End
   Begin VB.Label lblLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Check for Updates: this is release R6c."
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   63
      Tag             =   "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73521&lngWId=1"
      ToolTipText     =   "Click to visit home page on PSC"
      Top             =   4920
      Width           =   4635
   End
   Begin VB.Label lblProjMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   58
      Top             =   5760
      Width           =   1125
   End
   Begin VB.Label lblSynColours 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax Colours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   57
      Top             =   5760
      Width           =   1275
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface to User Configurable Options for this program."

' what?
'  user interface to provide various persistable options / preferences.
' why?
'  to add a couple of nice things.
' when?
'  as and when user chooses.
'  options are generally read as and when used.
' how?
'  user selects something that calls this form.
' who?
'  d.c.

' Notes:
' ... The following User Preferences are currently available:
'   ... User Configured location for VB IDEs and Text/Code Editor program.
'   ... Whether Show in #IDE / Open in Editor are available.
'   ... Viewer Colours for Syntax Colouring.
'   ... Viewer Font Name/Size.
'   ... Viewer Font Bold.
'   ... Viewer Line Numbers.
'   ... Location of Unzip Folder.
'   ... Auto-Unzip.
'   ... Auto-Load VBPs from UnZip.
'   ... Auto-Clean/Remove Unzip Folder on Program Exit.

Option Explicit

' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long
' -------------------------------------------------------------------

' ... sample code to display in the temp viewer.
Private mExample As String

' ... turns off/on processing interface related methods.
Private mLoading As Boolean


Private Sub pLoadColourCursors()

' ... sets up the hand cursor for the colour labels.

Dim lngCount As Long
Dim lngLoop As Long

    If myHand_handle = 0 Then Exit Sub
    
    lngCount = lblColour.Count
    For lngLoop = 0 To lngCount - 1
        lblColour(lngLoop).MouseIcon = myHandCursor
        lblColour(lngLoop).MousePointer = vbCustom
    Next lngLoop
    
End Sub

Private Sub pLoadLinkCursors()

' ... sets up the hand cursor for the link labels.

Dim lngCount As Long
Dim lngLoop As Long

    If myHand_handle = 0 Then Exit Sub
    
    lngCount = lblLink.Count
    For lngLoop = 0 To lngCount - 1
        lblLink(lngLoop).MouseIcon = myHandCursor
        lblLink(lngLoop).MousePointer = vbCustom
    Next lngLoop
    
End Sub

Private Sub pSetUpExample()

' ... some sample code for the viewer.

Dim x As SBuilder ' StringWorker

    Set x = New SBuilder ' StringWorker
    
    x.AppendAsLine "Private Function DoSomething(pThis As Long, pThat As String) As Long"
    x.AppendAsLine "Attribute DoSomething.VB_Description = " & Chr$(34) & "Does something or other." & Chr$(34)
    x.AppendAsLine "' ... does something or other :-)"
    x.AppendAsLine "Dim i As Long"
    x.AppendAsLine "    On Error Goto ErrHan:"
    x.AppendAsLine "    For i = 0 To pThis"
    x.AppendAsLine "        If CDbl(i) = Val(pThat) Then"
    x.AppendAsLine "            Debug.Print " & Chr$(34) & "Did Something" & Chr$(34)
    x.AppendAsLine "            DoSomething = 1"
    x.AppendAsLine "            Exit For"
    x.AppendAsLine "        End If"
    x.AppendAsLine "    Next i"
    x.AppendAsLine "Exit Function"
    x.AppendAsLine "ErrHan:"
    x.AppendAsLine "    "
    x.AppendAsLine "End Function ' ... DoSomething"
    
    mExample = x.TheString
    
    Set x = Nothing
    
End Sub

Private Sub pLoadHandCursor()

' ... try and load the hand cursor.

    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        
    End If

End Sub

Private Sub cboFont_Click(Index As Integer)
    chkSyntaxColours_Click
End Sub

Private Sub chkAttributes_Click()
    chkSyntaxColours_Click
End Sub

Private Sub chkBold_Click()
    chkSyntaxColours_Click
End Sub

Private Sub chkLineNos_Click()
    chkSyntaxColours_Click
End Sub

Private Sub chkShowMenu_Click(Index As Integer)

' ... enable / disable open dialog buttons and respective path labels.
Dim bEnabled As Boolean
    
    bEnabled = chkShowMenu(Index).Value = VBRUN.vbChecked
    cmdPath(Index).Enabled = bEnabled
    lblPath(Index).Enabled = bEnabled

End Sub

Private Sub chkSyntaxColours_Click()

' ... respond to Use Own Colours Check Box click.

Dim lngCount As Long
Dim lngLoop As Long
Dim bEnabled As Boolean
Dim sColourTbl As String

    If mLoading Then Exit Sub
    
    lngCount = lblSynColours.Count
    If lngCount = 0 Then Exit Sub
        
    bEnabled = chkSyntaxColours.Value = VBRUN.vbChecked
    ' ... enable / disable label controls dependent upon Check Box Value.
    For lngLoop = 1 To lngCount - 1
        lblSynColours(lngLoop).Enabled = bEnabled
        lblColour(lngLoop - 1).Enabled = bEnabled
    Next lngLoop
    
    
    If bEnabled Then
        ' ... if use own colours then build a colour table for the BuildRTFString method.
        sColourTbl = modEncode.BuildRTFColourTable(lblColour(1).BackColor, _
                                                   lblColour(2).BackColor, _
                                                   lblColour(3).BackColor, _
                                                   lblColour(4).BackColor, _
                                                   lblColour(5).BackColor)
        rtb.BackColor = lblColour(0).BackColor
    Else
        ' ... just resort to default values.
        rtb.BackColor = &HFFFFFF
    End If
    
    pSetRTF sColourTbl
    
    sColourTbl = vbNullString

End Sub

Private Sub chkTipOfTheDay_Click()
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkTipOfTheDay.Value

End Sub

Private Sub chkUseChildWindows_Click()
Dim bEnabled As Boolean
    bEnabled = chkUseChildWindows.Value And vbChecked
    chkHideProj.Enabled = bEnabled
    chkHideTool.Enabled = bEnabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Attribute cmdOK_Click.VB_Description = "Saves Current Option Values when clicked."

Dim xOptions As cOptions

    Set xOptions = New cOptions
    
    xOptions.Read
    
        ' ... syntax colours.
    xOptions.UseOwnColours = chkSyntaxColours.Value And vbChecked
    
    xOptions.ViewerBackColour = lblColour(0).BackColor
    xOptions.NormalTextColour = lblColour(1).BackColor
    xOptions.KeywordTextColour = lblColour(2).BackColor
    xOptions.CommentTextColour = lblColour(3).BackColor
    xOptions.AttributeTextColour = lblColour(4).BackColor
    xOptions.LineNoTextColour = lblColour(5).BackColor
    xOptions.LineNumbers = chkLineNos.Value And vbChecked
    
    xOptions.FontBold = chkBold.Value And vbChecked
    xOptions.ShowAttributes = chkAttributes.Value And vbChecked
        
    If cboFont(0).ListIndex > -1 Then xOptions.FontSize = cboFont(0).Text
    If cboFont(1).ListIndex > -1 Then xOptions.FontName = cboFont(1).Text
        
        ' ... general.
    xOptions.UseChildWindows = chkUseChildWindows.Value And vbChecked
    xOptions.HideChildProject = chkHideProj.Value And vbChecked
    xOptions.HideChildToolbar = chkHideTool.Value And vbChecked
    xOptions.ConfirmExit = chkConfirmExit.Value And vbChecked
    xOptions.ShowClassHeadCount = chkShowCHeadCount.Value And vbChecked
    xOptions.AutoRTFEncoding = chkCodeAutoEncode.Value And vbChecked
    xOptions.AutoLoadAllCode = chkCodeLoadAll.Value And vbChecked
        
        ' ... unzip
    xOptions.AutoUnZip = chkAutoUnzip.Value And vbChecked
    xOptions.AutoLoadProjects = chkAutoLoadVBP.Value And vbChecked
    xOptions.AutoCleanUnzipFolder = chkAutoCleanUnzipFolder.Value And vbChecked
    xOptions.UnzipFolder = lblUnzipFolder.Caption

        ' ... project menu.
    xOptions.ShowVB432IDE = chkShowMenu(0).Value = vbChecked
    xOptions.ShowVB5IDE = chkShowMenu(1).Value = vbChecked
    xOptions.ShowVB6IDE = chkShowMenu(2).Value = vbChecked ' (could this be And vbChecked?)
    xOptions.ShowTextEditor = chkShowMenu(3).Value And vbChecked ' let's try.
    
    
    If Len(lblPath(0).Caption) Then
        xOptions.PathToVB432 = lblPath(0).Caption
    End If
    If Len(lblPath(1).Caption) Then
        xOptions.PathToVB5 = lblPath(1).Caption
    End If
    If Len(lblPath(2).Caption) Then
        xOptions.PathToVB6 = lblPath(2).Caption
    End If
    If Len(lblPath(3).Caption) Then
        xOptions.PathToTextEditor = lblPath(3).Caption
    End If
    
    xOptions.Save
    
    Set xOptions = Nothing
    
    ' ... Unload the form.
    cmdCancel_Click
    
End Sub

Private Sub pSetRTF(Optional pColourTable As String = vbNullString)

' ... process and display text for rt box.

    rtb.TextRTF = modEncode.BuildRTFString(mExample, cboFont(1).Text, , cboFont(0).Text, pColourTable, IIf(chkLineNos.Value = vbChecked, True, False), IIf(chkAttributes.Value = vbChecked, True, False), IIf(chkBold.Value = vbChecked, True, False))
    
End Sub

Private Sub pInit()

' ... set up the form.

Dim lngCount As Long
Dim lngLoop As Long
Dim xOptions As cOptions

Dim lngFCount As Long
Dim sFont As String
Dim sUserFont As String
Dim sFontSize As String
Dim sUserFontSize As String
Dim lngCFontIndex As Long
Dim ShowAtStartup As Long

    mLoading = True
    ' -------------------------------------------------------------------
    ' ... v8.
    lblLink(5).Caption = "Check for updates, this is release " & AppVersion
    ' -------------------------------------------------------------------
    ' ... rich text box: turn word wrap off.
    modGeneral.WordWrapRTFBox rtb.hwnd
    
    pSetUpExample
    pLoadHandCursor
    pLoadColourCursors
    pLoadLinkCursors
    
    Set xOptions = New cOptions
    
    xOptions.Read
    
        ' ... syntax colours.
    lblColour(0).BackColor = xOptions.ViewerBackColour
    lblColour(1).BackColor = xOptions.NormalTextColour
    lblColour(2).BackColor = xOptions.KeywordTextColour
    lblColour(3).BackColor = xOptions.CommentTextColour
    lblColour(4).BackColor = xOptions.AttributeTextColour
    lblColour(5).BackColor = xOptions.LineNoTextColour
        
    chkLineNos.Value = IIf(xOptions.LineNumbers, vbChecked, vbUnchecked)
    chkSyntaxColours.Value = IIf(xOptions.UseOwnColours, vbChecked, vbUnchecked)
        
    sUserFont = xOptions.FontName
    sUserFontSize = xOptions.FontSize
    chkBold.Value = IIf(xOptions.FontBold, vbChecked, vbUnchecked)
    chkAttributes.Value = IIf(xOptions.ShowAttributes, vbChecked, vbUnchecked)
        
        ' ... project menu.
    lblPath(0).Caption = xOptions.PathToVB432
    lblPath(1).Caption = xOptions.PathToVB5
    lblPath(2).Caption = xOptions.PathToVB6
    lblPath(3).Caption = xOptions.PathToTextEditor
    
    chkShowMenu(0).Value = IIf(xOptions.ShowVB432IDE, vbChecked, vbUnchecked)
    chkShowMenu(1).Value = IIf(xOptions.ShowVB5IDE, vbChecked, vbUnchecked)
    chkShowMenu(2).Value = IIf(xOptions.ShowVB6IDE, vbChecked, vbUnchecked)
    chkShowMenu(3).Value = IIf(xOptions.ShowTextEditor, vbChecked, vbUnchecked)
    
        ' ... unzip.
    lblUnzipFolder.Caption = xOptions.UnzipFolder
    chkAutoUnzip.Value = IIf(xOptions.AutoUnZip, vbChecked, vbUnchecked)
    chkAutoLoadVBP.Value = IIf(xOptions.AutoLoadProjects, vbChecked, vbUnchecked)
    chkAutoCleanUnzipFolder.Value = IIf(xOptions.AutoCleanUnzipFolder, vbChecked, vbUnchecked)
    
    ' ... general.
    chkUseChildWindows.Value = IIf(xOptions.UseChildWindows, vbChecked, vbUnchecked)
    chkHideProj.Value = IIf(xOptions.HideChildProject, vbChecked, vbUnchecked)
    chkHideTool.Value = IIf(xOptions.HideChildToolbar, vbChecked, vbUnchecked)
    chkConfirmExit.Value = IIf(xOptions.ConfirmExit, vbChecked, vbUnchecked)
    chkShowCHeadCount.Value = IIf(xOptions.ShowClassHeadCount, vbChecked, vbUnchecked)
    chkCodeAutoEncode.Value = IIf(xOptions.AutoRTFEncoding, vbChecked, vbUnchecked)
    chkCodeLoadAll.Value = IIf(xOptions.AutoLoadAllCode, vbChecked, vbUnchecked)
    
    ' ... font sizes followed by names.
    lngCFontIndex = -1
    For lngLoop = 7 To 20
        sFontSize = Format$(lngLoop, "00")
        cboFont(0).AddItem sFontSize
        If sFontSize = sUserFontSize Then
            lngCFontIndex = cboFont(0).ListCount - 1
        End If
    Next lngLoop
    
    If lngCFontIndex > -1 Then
        cboFont(0).ListIndex = lngCFontIndex
    End If
    
    lngCFontIndex = -1
    lngFCount = VB.Screen.FontCount
    
    For lngLoop = 0 To lngFCount - 1
        sFont = VB.Screen.Fonts(lngLoop)
        If sFont = sUserFont Then
            lngCFontIndex = lngLoop
        End If
        cboFont(1).AddItem sFont
    Next lngLoop
    
    If lngCFontIndex > -1 Then
        cboFont(1).Text = sUserFont
    Else
        On Error Resume Next
        cboFont(1).Text = "Courier New" ' ... force a click if found courier new (default font) font.
        If Err.Number <> 0 Then
            Err.Clear
            cboFont(1).ListIndex = 0        ' ... first font if no courier.
        End If
    End If
    
    Set xOptions = Nothing
    
    lngCount = cmdPath.Count
    
    If lngCount = 0 Then Exit Sub
    
    For lngLoop = 0 To lngCount - 1
        chkShowMenu_Click CInt(lngLoop)
    Next lngLoop
    
    mLoading = False
    
    chkUseChildWindows_Click
    chkSyntaxColours_Click
    
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkTipOfTheDay.Value = ShowAtStartup
    
End Sub

Private Sub cmdPath_Click(Index As Integer)

' ... open dialog box to select program locations.

Dim sFileName As String
Dim sFilter As String
Dim sInitDir As String
Dim sDialogTitle As String
Dim sTmpTitle As String
Dim xOptions As cOptions

    On Error GoTo ErrHan:
    
    sTmpTitle = "Select Path to "
    Set xOptions = New cOptions
    xOptions.Read
    
    ' ... set up the open dialog.
    Select Case Index
    
        Case 0:
                
                sDialogTitle = "VB4 IDE"
                sFilter = modDialog.MakeDialogFilter("VB4 IDE", "VB32", "Exe")
                sInitDir = xOptions.PathToVB432
                sFileName = "VB32.Exe"
                
        Case 1:
                
                sDialogTitle = "VB5 IDE"
                sFilter = modDialog.MakeDialogFilter("VB5 IDE", "VB5", "Exe")
                sInitDir = xOptions.PathToVB5
                sFileName = "VB5.Exe"
                
        Case 2:
                
                sDialogTitle = "VB6 IDE"
                sFilter = modDialog.MakeDialogFilter("VB6 IDE", "VB6", "Exe")
                sInitDir = xOptions.PathToVB6
                sFileName = "VB6.Exe"
        Case 3:
                sDialogTitle = "Code Editor"
                sFilter = modDialog.MakeDialogFilter()
                sInitDir = xOptions.PathToTextEditor
                sFileName = ""
                
    End Select
    
    sDialogTitle = sTmpTitle & sDialogTitle
    ' ... show the open dialog.
    
    sFileName = modDialog.GetOpenFileName(sFileName, , sFilter, , sInitDir, sDialogTitle)
    If Len(sFileName) Then lblPath(Index).Caption = sFileName
    
ResumeError:
    On Error GoTo 0
    Set xOptions = Nothing
    
Exit Sub

ErrHan:

    Debug.Print "frmOptions.cmdPath_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:
    
End Sub

Private Sub cmdUnzipFolder_Click()

Dim sFolder As String
    
    sFolder = modDialog.GetFolder("Select Unzip Folder", App.Path) ' Me.hwnd)
    If Len(sFolder) Then
        If LCase$(sFolder) = LCase$(App.Path) Then
            ' ... v6, protect against unzipping to program folder
            ' ... else when killing unzip folder the program files
            ' ... will also be deleted.
            MsgBox "Unzip Folder is not allowed to be the same as the Program's folder.", vbInformation, Caption
            sFolder = sFolder & "\" & c_def_UnZipFolder
        End If
        lblUnzipFolder.Caption = sFolder
    End If
    
    sFolder = vbNullString
        
End Sub

Private Sub Form_Load()

' ... set up.
    
    pInit
    ' ... flags for Choose Colour and Open FIle dialogs.
    cdgColour.flags = &H2 Or &H1                ' ... open full and set colour (MSComDlg.ColorConstants)
                                                ' ... _ (MSComDlg.FileOpenConstants)
    ClearMemory
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearMemory
End Sub

Private Sub lblColour_Click(Index As Integer)

' ... select respective colour from color dialog
' ... and re-process viewer text.

Dim lngColour As Long
Dim lngCursor As StdPicture
Dim lngPointer As Long

    On Error GoTo ErrHan
    
    Set lngCursor = lblColour(Index).MouseIcon
    lngPointer = lblColour(Index).MousePointer
    
    lblColour(Index).MousePointer = vbNormal
    
    cdgColour.Color = lblColour(Index).BackColor    ' ... set the colour of the dialog.
    cdgColour.ShowColor                             ' ... load the color dialog.
        
    lngColour = cdgColour.Color
    lblColour(Index).BackColor = lngColour
    
ResumeError:

    lblColour(Index).MousePointer = lngPointer
    lblColour(Index).MouseIcon = lngCursor
    
    ' ... force re-process of viewer text.
    chkSyntaxColours_Click
    
Exit Sub

ErrHan:

    Debug.Print "frmOptions.lblColour_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub lblLink_Click(Index As Integer)
    modGeneral.OpenWebPage lblLink(Index).Tag
End Sub

Private Sub TabStrip1_Click()

' ... show / hide appropriate picture boxes dependent upon current tab.

Dim lngSelectedTab As Long
Dim lngCount As Long
Dim lngLoop As Long
    
    lngCount = TabStrip1.Tabs.Count
    If lngCount = 0 Then Exit Sub
    For lngLoop = 1 To lngCount
        If TabStrip1.Tabs(lngLoop).Selected = True Then
            lngSelectedTab = lngLoop
            Exit For
        End If
    Next lngLoop
    
    lngCount = picOpt.Count
    For lngLoop = 0 To lngCount - 1
        picOpt(lngLoop).Visible = False
    Next lngLoop
    
    If lngSelectedTab > 0 Then
        picOpt(lngSelectedTab - 1).Visible = True
        picOpt(lngSelectedTab - 1).ZOrder
    End If

End Sub
