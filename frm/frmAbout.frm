VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4875
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5070
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1965
      TabIndex        =   6
      Top             =   4410
      Width           =   1260
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "myspace"
      ForeColor       =   &H00808080&
      Height          =   270
      Index           =   4
      Left            =   210
      TabIndex        =   9
      Tag             =   "http://www.myspace.com/dave_e_c"
      ToolTipText     =   "dave's guitar tunes on myspace"
      Top             =   4020
      Width           =   4635
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Oct 2010 - Mar 2011"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   210
      TabIndex        =   8
      Top             =   1140
      Width           =   4635
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "enjoy"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   3360
      Width           =   4635
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "(mostly) Written by Dave Carter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   780
      Width           =   4635
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmAbout.frx":058A
      ForeColor       =   &H00000000&
      Height          =   1110
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   2190
      Width           =   4635
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   210
      TabIndex        =   0
      Top             =   1620
      Width           =   4635
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4635
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   540
      Width           =   4635
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   3720
      Width           =   4635
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "About Form"
Option Explicit

' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long

Private Sub cmdOK_Click()
Attribute cmdOK_Click.VB_Description = "unloads the About Form."
  Unload Me
End Sub

Private Sub Form_Load()
Attribute Form_Load.VB_Description = "sets up the About Form Info."
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.FileDescription
    lblDisclaimer.Caption = App.LegalCopyright
    pLoadLinkCursor
End Sub

Private Sub Form_LostFocus()
Attribute Form_LostFocus.VB_Description = "unloads the About Form when it has lost focus."
    cmdOK_Click
End Sub

Private Sub lblInfo_Click(Index As Integer)
Attribute lblInfo_Click.VB_Description = "links to the author's page on myspace."
    If Index = 4 Then
        modGeneral.OpenWebPage lblInfo(Index).Tag
    End If
End Sub

Private Sub pLoadHandCursor()
Attribute pLoadHandCursor.VB_Description = "attempts to load the hand cursor for the link to myspace label."

' ... try and load the hand cursor.

    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        
    End If

End Sub

Private Sub pLoadLinkCursor()
Attribute pLoadLinkCursor.VB_Description = "attempts to set the hand cursor for the myspace link label."

' ... sets up the hand cursor for the link labels.

    pLoadHandCursor
    
    If myHand_handle = 0 Then Exit Sub
    
    lblInfo(4).MouseIcon = myHandCursor
    lblInfo(4).MousePointer = vbCustom
    
End Sub

