VERSION 5.00
Begin VB.Form frmPSCReadMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Read Me Text"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   26
   Icon            =   "frmPSCReadMe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmPSCReadMe.frx":058A
      Top             =   1380
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "This file came from Planet-Source-Code.com...the home millions of lines of source code"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   5400
      Width           =   6375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      Caption         =   "Link to PSC Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   4800
      Width           =   6375
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Submission Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   6315
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPSCReadMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface to viewing a PSC Read Me File attached to a downloaded project from planet source code."
Option Explicit

Private mPSCText As String
Private mPSCName As String
Private mPSCDesc As String
Private mPSCLink As String

' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long

Public Sub LoadReadMe(pvReadMeText As String)
Dim lngFound As Long
Dim lngNext As Long
Dim lngStart As Long
Dim sFind As String
'Dim sTmp As String

    pInit
    
    mPSCText = pvReadMeText
    mPSCDesc = mPSCText
    lngStart = 1
    sFind = ": "
    
    lngFound = InStr(lngStart, mPSCText, sFind)
    
    If lngFound > 0 Then
        lngStart = lngFound + Len(sFind)
        sFind = vbCrLf
        lngNext = InStr(lngStart, mPSCText, sFind)
        
        If lngNext > 0 Then
            mPSCName = Mid$(mPSCText, lngStart, lngNext - lngStart)
            lngStart = lngNext + 1
            sFind = ": "
            lngFound = InStr(lngStart, mPSCText, sFind)
            If lngFound > 0 Then
                lngStart = lngFound + Len(sFind)
                sFind = "This file came from Planet-Source-Code.com...the home millions of lines of source code"
                lngNext = InStr(lngStart, mPSCText, sFind)
                If lngNext > 0 Then
                    mPSCDesc = Mid$(mPSCText, lngStart, lngNext - lngStart)
                    lngStart = lngNext + 1
                    sFind = "You can view comments on this code/and or vote on it at:"
                    lngFound = InStr(lngStart, mPSCText, sFind)
                    If lngFound > 0 Then
                        lngStart = lngFound + Len(sFind)
                        sFind = vbCrLf
                        lngNext = InStr(lngStart, mPSCText, sFind)
                        If lngNext > 0 Then
                            mPSCLink = Mid$(mPSCText, lngStart, lngNext - lngStart)
                            mPSCLink = Trim$(mPSCLink)
                        End If
                    End If
                End If
            End If
        End If
        
    End If
    
    lblName.Caption = mPSCName
    
    ' ... replace char 10 with new line as intended.
    txtDesc.Text = modStrings.Replace(mPSCDesc, Chr$(10), vbCrLf)
    
    lblLink.Caption = "Goto PSC Page" '
    lblLink.Tag = mPSCLink
    
End Sub

Private Sub pInit()

    mPSCText = vbNullString
    mPSCName = vbNullString
    mPSCDesc = vbNullString
    mPSCLink = vbNullString
    
    lblName.Caption = ""
    txtDesc.Text = ""
    
    lblLink.Caption = ""
    lblLink.Tag = ""
    
    pLoadHandCursor
    
End Sub

Private Sub pRelease()
    pInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pRelease
    ClearMemory
End Sub

Private Sub lblLink_Click()
    If Len(lblLink.Tag) Then
        modGeneral.OpenWebPage lblLink.Tag
    End If
End Sub

Private Sub pLoadHandCursor()

' ... try and load the hand cursor.

    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        If myHand_handle = 0 Then Exit Sub
        
        lblLink.MouseIcon = myHandCursor
        lblLink.MousePointer = vbCustom
        
    End If

End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblLink.ForeColor = vbBlue ' vbBlack
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblLink.ForeColor = &H800000
End Sub
