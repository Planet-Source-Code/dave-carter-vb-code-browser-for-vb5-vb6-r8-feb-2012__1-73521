VERSION 5.00
Begin VB.Form frmMngVBPHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage VBP History"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   25
   Icon            =   "frmMngVBPHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox lstHistory 
      Height          =   2085
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblSelCount 
      Caption         =   "Selected Count: "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Tag             =   "Selected Count: "
      Top             =   3520
      Width           =   2895
   End
   Begin VB.Label lblItemCount 
      Caption         =   "Item Count: "
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Tag             =   "Item Count: "
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Select (check) VBP Items you would like to keep and click the OK button else just click Cancel to exit this form."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmMngVBPHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Form to manage VBP History"

' what?
'  user interface to manage items in vbp history.
' why?
'  make it easier to do manage the number of items listed.
' when?
'  as and when user chooses.
' how?
'  user checks items wanted and unchecks items no longer wanted and
'  then clicks ok button.
'  clicking the cancel button will close the form with no
'  changes being saved.
' who?
'  d.c.

Option Explicit

Private moFileHistory As StringArray
Private Const cHistFileName As String = "History.dat"
Private mLoadingData As Boolean

Private Sub pLoadHistList()
' ... load history items into list box and check each one.
Dim sFile As String
Dim lngCount As Long
Dim lngLoop As Long
Dim sItem As String
Dim lngNewIndex As Integer

    ' ... clear out any existing hist menu items.
    lstHistory.Clear
    
    ' ... set up history string array.
    pSetNewHistory
    
    ' ... read the history file for the items to load.
    sFile = App.Path & "\" & cHistFileName
    
    If Dir$(sFile, vbNormal) <> "" Then
        
        moFileHistory.FromFile sFile, vbCrLf
        
        lngCount = moFileHistory.Count
        
        If lngCount > 0 Then
        
            mLoadingData = True
            
            ' ... load the items found in the history file.
            For lngLoop = 1 To lngCount
                
                sItem = moFileHistory(lngLoop)
                
                lstHistory.AddItem sItem
                
                lngNewIndex = lstHistory.NewIndex
                lstHistory.Selected(lngNewIndex) = True
            
            Next lngLoop
            
            lstHistory.ListIndex = -1
            
            mLoadingData = False
            
        End If
    
    End If
    
End Sub

Private Sub pReleaseHistory()
' ... release any existing history string array.
    If Not moFileHistory Is Nothing Then
        moFileHistory.Clear
        Set moFileHistory = Nothing
    End If
End Sub

Private Sub pSetNewHistory()
' ... set up a new history string array.
    pReleaseHistory
    Set moFileHistory = New StringArray
    moFileHistory.DuplicatesAllowed = False
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
' ... save user preferences and exit.
Dim sFile As String
Dim lngCount As Long
Dim lngSelected As Long
Dim lngLoop As Long
Dim sItem As String

    pSetNewHistory
    
    ' ... set up history dat file name.
    sFile = App.Path & "\" & cHistFileName
    
    lngCount = lstHistory.ListCount
    If lngCount > 0 Then
        For lngLoop = 0 To lngCount - 1
            If lstHistory.Selected(lngLoop) = True Then
                lngSelected = lngSelected + 1
            End If
        Next lngLoop
        If lngSelected > 0 Then
            For lngLoop = 0 To lngCount - 1
                If lstHistory.Selected(lngLoop) = True Then
                    sItem = lstHistory.List(lngLoop)
                    moFileHistory.AddItemString sItem
                End If
            Next lngLoop
            ' ... save the history file.
            moFileHistory.ToFile sFile, vbCrLf
        Else
            ' ... no items so delete the history file.
            ' ... had to introduce this because StringArray wouldn't
            ' ... save an empty file.
            Kill sFile
        End If
    End If
    
    btnCancel_Click
    
End Sub

Private Sub Form_Load()
' ... entry, load history and show info.
    pLoadHistList
    pUpdateTotals
    btnOK.Enabled = lstHistory.ListCount > 0
    ClearMemory
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
' ... unloading so release history string array.
    pReleaseHistory
    ClearMemory
End Sub

Private Sub lstHistory_ItemCheck(Item As Integer)
' ... respond to user item check/uncheck.
    If mLoadingData Then Exit Sub
    pUpdateTotals

End Sub

Private Sub pUpdateTotals()
' ... reprint the totals to the count labels.
Dim lngSelected As Long
Dim lngCount As Long
Dim lngLoop As Long

    lngCount = lstHistory.ListCount
    
    If lngCount > 0 Then
    
        For lngLoop = 0 To lngCount - 1
        
            If lstHistory.Selected(lngLoop) = True Then
                lngSelected = lngSelected + 1
            End If
        
        Next lngLoop
    
    End If
    
    ' ... note: the base caption is written to the label's tag property.
    lblItemCount.Caption = lblItemCount.Tag & lngCount
    lblSelCount.Caption = lblSelCount.Tag & lngSelected
    
End Sub
