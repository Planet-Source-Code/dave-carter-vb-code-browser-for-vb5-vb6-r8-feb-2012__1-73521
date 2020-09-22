VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmUnZip 
   Caption         =   "UnZip Project"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   6
   Icon            =   "frmUnZip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCanvas 
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   6855
      TabIndex        =   2
      Top             =   1320
      Width           =   6855
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4320
         TabIndex        =   9
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelZip 
         Caption         =   "Select Zip"
         Height          =   495
         Left            =   60
         TabIndex        =   4
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdExtract 
         Caption         =   "UnZip"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5580
         TabIndex        =   3
         Top             =   5280
         Width           =   1215
      End
      Begin ComctlLib.ListView lv 
         Height          =   3795
         Left            =   60
         TabIndex        =   5
         Top             =   540
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6694
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         _Version        =   327682
         SmallIcons      =   "ilZip"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Label lblUnzipSize 
         Caption         =   "Unzip Size..."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1380
         TabIndex        =   13
         Top             =   5580
         Width           =   2835
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgBtn 
         Height          =   300
         Index           =   1
         Left            =   6480
         Picture         =   "frmUnZip.frx":014A
         Stretch         =   -1  'True
         ToolTipText     =   "PSC Read Me"
         Top             =   60
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblStats 
         Caption         =   "Zip Stats..."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1380
         TabIndex        =   8
         Top             =   5340
         Width           =   2835
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUnzipFile 
         Caption         =   "Unzip File:"
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   90
         TabIndex        =   7
         Top             =   4620
         Width           =   6735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFolder 
         Caption         =   "Unzip Folder:"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   6735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picFileView 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   5955
      ScaleWidth      =   6855
      TabIndex        =   10
      Top             =   1320
      Width           =   6915
      Begin RichTextLib.RichTextBox rtb 
         Height          =   5355
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   9446
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmUnZip.frx":06D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblZipFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip FIle Name:"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgBtn 
         Height          =   240
         Index           =   0
         Left            =   6540
         Picture         =   "frmUnZip.frx":0754
         ToolTipText     =   "Back"
         Top             =   120
         Width           =   240
      End
   End
   Begin ComctlLib.ImageList ilZip 
      Left            =   6420
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnZip.frx":0CDE
            Key             =   "DEFAULT"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnZip.frx":0DF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      Caption         =   "Unzip with VB by Chris Eastwood"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "http://www.codeguru.com/vb/gen/vb_graphics/fileformats/article.php/c6743"
      Top             =   960
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Label lblNoUnzip 
      Alignment       =   2  'Center
      Caption         =   $"frmUnZip.frx":1342
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   6915
   End
End
Attribute VB_Name = "frmUnZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "GUI to Unzip Function."
' what?
'  a simple user interface to the unzip funtion.
' why?
'  user control over unzipping.
' when?
'  you want to see the contents of a zip file and or unzip them.
' how?
'  a zip file name is required, this may be passed to the form via
'  frmUnzip.ReadZip sZipFileName
'  or
'  the user selects a zip file via the open file name dialog having clicked ' Select Zip '.
'  which will load the list view with the contents of the zip file.
'  To extract the zip file to the UnZip folder (see Options) the
'  'Extract ' button should be clicked.
'  If the extraction was successful the inzip form will then unload.

Option Explicit

' v8
Private mUnCompSize As Long                 ' uncompressed size of data

' ... variables for the Hand Cursor.
Private myHandCursor As StdPicture
Private myHand_handle As Long

Private mTheZipFile As String
Private mLMouseDown As Boolean

'Private Const LVM_FIRST As Long = &H1000
'Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
'Private Const LVS_EX_FULLROWSELECT As Long = &H20
'Private Const LVS_EX_CHECKBOXES As Long = &H4

'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mUnzippedFiles As String
Private mReadZipFiles As String
Private mVBPFileString As String
Private mVBGFileString As String
Private mPSCFile As String
Private mPSCFileText As String
Private mNoOfVBPs As Long
Private mNoOfVBGs As Long
Private mNoOfFilesInZip As Long
Private mNoOfUnzippedFiles As Long
Private mtFileInfo As FileNameInfo
Private mtPSCInfo As PSCInfo
Private mUnzipFolder As String
Private mCancelled As Boolean

Private mLastSourcePath As String


Public Property Get PSCText() As String
    PSCText = mPSCFileText
End Property

Public Property Get Cancelled() As Boolean
Attribute Cancelled.VB_Description = "Returns a Boolean describing whether the user clicked the cancel button from the Unzip form."
    Cancelled = mCancelled
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Reset everything internally for new use, call this before Reading."
    pInit
End Sub

Private Sub cmdCancel_Click()
    mCancelled = True
    Unload Me
End Sub

Private Sub cmdExtract_Click()
Dim bOK As Boolean
    
    Screen.MousePointer = vbHourglass
    
    cmdExtract.Enabled = False
    Me.Enabled = False
    
    bOK = pExtract(mTheZipFile)
    
    cmdExtract.Enabled = True
    Me.Enabled = True
    
    ClearMemory
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSelZip_Click()

Dim sDLGFilter As String
Dim sFileName As String
Dim xFileInfo As FileNameInfo

    On Error GoTo ErrHan:
    
    Screen.MousePointer = vbHourglass
    
    cmdSelZip.Enabled = False
    Me.Enabled = False
    
    sDLGFilter = modDialog.MakeDialogFilter("Zip Files", , "zip")
    sFileName = modDialog.GetOpenFileName(, , sDLGFilter, , mLastSourcePath, "Select Zip File", , Me.hwnd)
    
    If Len(sFileName) Then
        
        modFileName.ParseFileNameEx sFileName, xFileInfo
        mLastSourcePath = xFileInfo.Path
        
        lblUnzipFile.Caption = "Unzip File: " & sFileName
        lblUnzipFile.Refresh
        
        pOpen sFileName, True
        
    End If

ResumeError:
    
    cmdSelZip.Enabled = True
    Me.Enabled = True
    
    Screen.MousePointer = vbDefault
    
Exit Sub

ErrHan:

    Debug.Print "frmUnZip.cmdSelZip_Click.Error: " & Err.Number & "; " & Err.Description

    Resume ResumeError:


End Sub

Public Property Get CountOfFilesInZip() As Long
Attribute CountOfFilesInZip.VB_Description = "Returns the no. of files found in the Zip File."
    CountOfFilesInZip = mNoOfFilesInZip
End Property

Public Property Get CountOfUnzippedFiles() As Long
Attribute CountOfUnzippedFiles.VB_Description = "Returns the no. of files that were (previously) Unzipped."
    CountOfUnzippedFiles = mNoOfUnzippedFiles
End Property

Public Property Get CountOfVBGs() As Long
Attribute CountOfVBGs.VB_Description = "Returns the no. of VBGs found in the Zip File."
    CountOfVBGs = mNoOfVBGs
End Property

Public Property Get CountOfVBPs() As Long
Attribute CountOfVBPs.VB_Description = "Returns the no. of VBPs found in the Zip File."
    CountOfVBPs = mNoOfVBPs
End Property

Private Sub Form_Load()
Dim xOptions As cOptions
Dim bHaveUnzip As Boolean

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    ' -------------------------------------------------------------------
    Set xOptions = New cOptions
    xOptions.Read
    lblFolder.Caption = "Unzip Folder: " & xOptions.UnzipFolder
    Set xOptions = Nothing
   ' -------------------------------------------------------------------
   
    pLoadLVColumns
    
    pLoadHandCursor
    
    pLoadButtonCursors
    
    ' -------------------------------------------------------------------
    ' ... check if unzip32 exists else inform user.
    bHaveUnzip = modGeneral.CheckUnzipAPI
    modGeneral.WordWrapRTFBox rtb.hwnd
    
    lblNoUnzip.Visible = bHaveUnzip = False
    lblLink.Visible = bHaveUnzip = False
    picCanvas.Visible = bHaveUnzip
    If bHaveUnzip Then
        picCanvas.Move 120, 180
        Height = (picCanvas.Top * 4) + picCanvas.Height
    Else
        Height = (lblLink.Top) + (lblLink.Height * 4)
    End If
    
    With picCanvas
        picFileView.Move .Left, .Top, .Width, .Height
        .ZOrder
    End With
    
    lblStats.Caption = ""
    lblUnzipSize.Caption = "" ' v8
    
    ClearMemory
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' ... if end form then set cancelled.
    If UnloadMode = 0 Then
        mCancelled = True
    End If
    If mCancelled = True Then
        mNoOfFilesInZip = lv.ListItems.Count
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearMemory
End Sub

Private Sub imgBtn_Click(Index As Integer)
    If Index = 0 Then
        picCanvas.ZOrder
    Else
        If Index = 1 Then
            If Len(mPSCFileText) Then
                frmPSCReadMe.LoadReadMe mPSCFileText
                frmPSCReadMe.Show vbModal
            End If
        End If
    End If
End Sub

Private Sub lblLink_Click()
    OpenWebPage lblLink.Tag
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLink.ForeColor = vbBlue
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLink.ForeColor = vbBlack
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

Dim lngSortIndex As Long

    ' ... reverse current sort order.
    lv.SortOrder = IIf(lv.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    
    lngSortIndex = ColumnHeader.Index
    
    Select Case lngSortIndex
        ' hidden:   7 = size, 8 = date, 9 = packed
        ' visible:  3 = size, 4 = date, 5 = packed
        Case 3, 4, 5
            ' ... defer nnumeric/date fields to their
            ' ... string counterparts for sorting.
            lngSortIndex = lngSortIndex + 4
    End Select
    
    lv.SortKey = lngSortIndex - 1
    lv.Sorted = True
    
End Sub

Private Sub lv_DblClick()

Dim itmx As ListItem
Dim sTmp As String
Dim xFileInfo As FileNameInfo

    lblZipFile.Caption = vbNullString
    
    If lv.ListItems.Count = 0 Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rtb.LoadFile "" ' ... clear the rtbox.
    Set itmx = lv.SelectedItem
    If Not itmx Is Nothing Then
        If Len(mTheZipFile) Then
            modUnZip.UnzipToMemory mTheZipFile, itmx.Tag, sTmp
            If Len(sTmp) Then
                Select Case LCase$(Right$(itmx.Tag, 3))
                    Case "frm", "cls", "bas", "ctl", "ppg", "dsr"
                        rtb.TextRTF = modEncode.BuildRTFString(sTmp, , , "9", , , True)
                    Case Else
                        rtb.Text = sTmp
                End Select
                modFileName.ParseFileNameEx mTheZipFile, xFileInfo
                lblZipFile.Caption = xFileInfo.File & " ..\" & itmx.Tag
                picFileView.ZOrder
            End If
        End If
    End If
    
    If Not itmx Is Nothing Then
        Set itmx = Nothing
    End If
    sTmp = vbNullString
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub lv_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim sFile As String
Dim sExt As String

    If Data.GetFormat(vbCFFiles) Then
        ' ... if it's an ole file...
        sFile = Data.Files(1)
        If Dir$(sFile, vbNormal) <> "" Then
            sExt = Right$(LCase$(sFile), 4)
            If sExt = ".zip" Then
                ' ... if it's a zip then load it.
                pOpen sFile, True
            End If
        End If
    End If
    
End Sub

Private Function pExtract(pTheZipFile As String) As Boolean

Dim sFolder As String
Dim lngRet As Long
Dim xFileInfo As FileNameInfo
Dim xOptions As cOptions
'Dim lngCount As Long

    On Error GoTo ErrHan:
    
    If Len(Trim$(pTheZipFile)) = 0 Then
        Err.Raise vbObjectError + 1000, , "No File Name specified for UnZip operation."
    End If
    
    Set xOptions = New cOptions
    xOptions.Read
    sFolder = xOptions.UnzipFolder
    
    If Dir(sFolder, vbDirectory) = "" Then
        MkDir sFolder
    End If
    
    modFileName.ParseFileNameEx pTheZipFile, xFileInfo
    
    sFolder = sFolder & "\" & xFileInfo.FileName
    If Dir(sFolder, vbDirectory) = "" Then
        MkDir sFolder
    End If
    
    lngRet = modUnZip.Unzip(pTheZipFile, sFolder, mUnzippedFiles, mNoOfUnzippedFiles)
    
ResumeError:

    On Error GoTo 0
    
    If lngRet <> 0 Then
        ' ... unzip failed.
        mNoOfFilesInZip = 0
    Else
        mtPSCInfo.UnzipFolder = sFolder
        mUnzipFolder = sFolder
        pWriteUnzipPSCDetails
        MsgBox "Files Unzipped: " & Format$(mNoOfUnzippedFiles, cNumFormat) & vbNewLine & "to " & sFolder, vbInformation, "Unzip"
    End If
    
    Set xOptions = Nothing
    
    If lngRet = 0 Then Unload Me

Exit Function

ErrHan:
    ' ... attempt to guard against file system errors
    Debug.Print "frmUnZip.pExtract.Error: " & Err.Number & "; " & Err.Description
    lngRet = Err.Number
    Resume ResumeError:
    
End Function

Private Sub pInit()
Dim xInfo As FileNameInfo
    
    ' -------------------------------------------------------------------
    lblStats.Caption = ""
    ' -------------------------------------------------------------------
    If lv.ListItems.Count > 0 Then
        lv.ListItems(1).Selected = True
        lv.ListItems.Clear
    End If
    ' -------------------------------------------------------------------
    ' ... reset things for new use.
    mUnzipFolder = vbNullString
    mNoOfVBPs = 0
    mNoOfVBGs = 0
    mNoOfFilesInZip = 0
    mNoOfUnzippedFiles = 0
    mVBPFileString = vbNullString
    mVBGFileString = vbNullString
    mPSCFile = vbNullString
    mPSCFileText = vbNullString
    mTheZipFile = vbNullString
    mUnzippedFiles = vbNullString
    mReadZipFiles = vbNullString
    mCancelled = False
    ' -------------------------------------------------------------------
    ' ... clear existing zip file info.
    With xInfo
        mtFileInfo.Extension = .Extension
        mtFileInfo.File = .File
        mtFileInfo.FileName = .FileName
        mtFileInfo.Path = .Path
        mtFileInfo.PathAndName = .PathAndName
    End With
    ' -------------------------------------------------------------------
    imgBtn(1).Visible = False ' ... v8, added psc read me button.
    ' -------------------------------------------------------------------
    mUnCompSize = 0 ' v8, added
    ' v9-----------------------------------------------------------------
    With mtPSCInfo
        .Name = vbNullString
        .Link = vbNullString
        .Description = vbNullString
    End With
    ' -------------------------------------------------------------------
End Sub

Private Sub pLoadHandCursor()

' ... try and load the hand cursor.

    myHand_handle = modHandCursor.LoadHandCursor
    
    If myHand_handle <> 0 Then
        
        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
        lblLink.MouseIcon = myHandCursor
        lblLink.MousePointer = vbCustom
        
    End If

End Sub

Private Sub pLoadLVColumns()
Dim xCol As ColumnHeader
    
    On Error GoTo ErrHan:
    
    If lv.ColumnHeaders.Count = 0 Then
    
        With lv
            With .ColumnHeaders
                
                Set xCol = .Add(, , "Filename", 160 * Screen.TwipsPerPixelX)
                Set xCol = .Add(, , "Type", 160 * Screen.TwipsPerPixelX)
                Set xCol = .Add(, , "Size", 32 * Screen.TwipsPerPixelX, 1)
                Set xCol = .Add(, , "Date", 96 * Screen.TwipsPerPixelX, 2)
                Set xCol = .Add(, , "Packed", 32 * Screen.TwipsPerPixelX, 1)
                Set xCol = .Add(, , "Folder", 160 * Screen.TwipsPerPixelX)
                
                Set xCol = .Add(, "Size", , 0)
                Set xCol = .Add(, "Date", , 0)
                Set xCol = .Add(, "Packed", , 0)
            
            End With
        End With
        
        modGeneral.LVFullRowSelect lv.hwnd
    
    End If

ResErr:
    On Error GoTo 0
    If Not xCol Is Nothing Then Set xCol = Nothing
    
Exit Sub

ErrHan:

    Debug.Print "frmUnZip.pLoadLVColumns.Error: " & Err.Number & "; " & Err.Description
    GoTo ResErr:
    
End Sub

Private Function pOpen(ByVal pvTheZipFile As String, Optional ByVal pvLoadListView As Boolean = False) As Boolean

Dim bOK As Boolean
Dim sIcon As String
Dim itmx As ListItem
Dim sZipFileInfo As String
Dim sa As StringArray
Dim lngCount As Long
Dim lngLoop As Long
Dim tZipInf As ZipMemberInfo
Dim sZipMember As String
Dim sFileInfo As String
Dim sTmpVBP As String
Dim sTmpExt As String
'Dim sTmpMemStr As String
Dim iScreenPointer As Long
' v9-----------------------------------------------------------------
'Dim xPSCInfo As PSCInfo
    
    On Error Resume Next    ' ... just for now and just because ...
    
    iScreenPointer = Screen.MousePointer
    
    pInit
    
    ' -------------------------------------------------------------------
    ' ... read the zip file contents, into sZipFileInfo.
    bOK = modUnZip.ReadZip(pvTheZipFile, sZipFileInfo, lngCount)
    
'    If bOK = True Then
'        ' ... use ReadZip above to validate zip file and existence.
        mTheZipFile = pvTheZipFile
        modFileName.ParseFileNameEx mTheZipFile, mtFileInfo
'
'    End If
    
    If Len(sZipFileInfo) = 0 Then Exit Function
    
    If pvLoadListView = True Then
        ' ... ensure columns available in list view.
        pLoadLVColumns
    
    End If
    
    
    Set sa = New StringArray
    sa.FromString sZipFileInfo, vbNewLine
    
    lngCount = sa.Count
    ' -------------------------------------------------------------------
    ' ... update total no. of files in zip.
    mNoOfFilesInZip = lngCount
    
    If lngCount > 0 Then
        
        pOpen = True
        
        For lngLoop = 1 To lngCount
            
            modGeneral.ParseZipMemberInfoItem sa, lngLoop, tZipInf
            
            With tZipInf
            
                ' v8, increment uncompressed size -----------------------------------
                mUnCompSize = mUnCompSize + .UnCompSize
                
                sZipMember = .FileName

                ' ... remember vbp files found.
                sTmpExt = LCase$(Right$(.FullPathAndName, 4))
                Select Case sTmpExt
                    Case ".vbp", ".vbg"
                        sTmpVBP = mtFileInfo.FileName & "\" & .FullPathAndName
                        If sTmpExt = ".vbp" Then
                            If Len(mVBPFileString) Then sTmpVBP = vbCrLf & sTmpVBP
                            mVBPFileString = mVBPFileString & sTmpVBP
                            mNoOfVBPs = mNoOfVBPs + 1
                        ElseIf sTmpExt = ".vbg" Then
                            If Len(mVBGFileString) Then sTmpVBP = vbCrLf & sTmpVBP
                            mVBGFileString = mVBGFileString & sTmpVBP
                            mNoOfVBGs = mNoOfVBGs + 1
                            
                        End If
                    Case ".txt"
                        If Left$(.FileName, 12) = "@PSC_ReadMe_" Then
                            mPSCFile = mtFileInfo.FileName & "\" & .FullPathAndName
                            modUnZip.UnzipToMemory mTheZipFile, .FullPathAndName, mPSCFileText
                        End If
                End Select
                
                If pvLoadListView Then
                
                    sIcon = AddIconToImageList(sZipMember, ilZip, "DEFAULT", sFileInfo)
    
                    If .Encrypted Then sZipMember = "+" & sZipMember
    
                    Set itmx = lv.ListItems.Add(, lngLoop & "File", sZipMember, , sIcon)
                    itmx.Tag = .FullPathAndName
                    
                    itmx.SubItems(1) = sFileInfo
                    itmx.SubItems(2) = Format$(.UnCompSize, cNumFormat)
                    itmx.SubItems(3) = Format$(.FileDate, "short date") & " " & Format$(.FileDate, "short time")
                    itmx.SubItems(4) = Format$(.CompSize, cNumFormat)
                    itmx.SubItems(5) = .FilePath
    
                    ' ... adding these for sorting on size, date and packed respectively,
                    ' ... when column header is clicked for one of these the sorting
                    ' ... will defer to its counterpart below.
    
                    itmx.SubItems(6) = Format$(.UnCompSize, "0000000000")
                    itmx.SubItems(7) = Format$(.FileDate, "0000000000.0000000000")
                    itmx.SubItems(8) = Format$(.CompSize, "0000000000")
                
                End If
                
            End With
        
        Next lngLoop
    
    End If
    
    If Not sa Is Nothing Then
        
        Set sa = Nothing
        
    End If
    
    ' v6, autosize columns.
    ' ... delay until updated auto size method with columns we want resized.
'''    If pvLoadListView Then
'''        If lv.ListItems.Count > 0 Then
'''            modGeneral.AutosizeColumns lv
'''        End If
'''    End If
    
    cmdExtract.Enabled = lv.ListItems.Count > 0
    
    lblStats.Caption = "No. Of Files: " & Format$(mNoOfFilesInZip, cNumFormat) & ".  No. of VBPs: " & Format$(mNoOfVBPs, cNumFormat)
    ' v8 ----------------------------------------------------------------
    lblUnzipSize.Caption = "Uncompressed: " & Format$(mUnCompSize, cNumFormat) & " bytes"
    ' -------------------------------------------------------------------
    If Len(mPSCFileText) Then
        ' -------------------------------------------------------------------
        ' v9
        ParsePSCInfo mPSCFileText, mtPSCInfo
'        With mtPSCInfo
'            MsgBox .Name & vbNewLine & .Link & vbNewLine & .Description
'        End With
        
        imgBtn(1).Visible = True ' ... v8, added psc read me button.
        ' -------------------------------------------------------------------
'        If iScreenPointer = vbHourglass Then
'            Screen.MousePointer = vbDefault
'        End If
'        If MsgBox("There is a PSC Read Me File available in this Zip;" & vbNewLine & "Would you like to view it now?", vbQuestion + vbYesNo, "PSC Read Me") = vbYes Then
'            frmPSCReadMe.LoadReadMe mPSCFileText
'            frmPSCReadMe.Show vbModal
'        End If
'        If iScreenPointer = vbHourglass Then
'            Screen.MousePointer = iScreenPointer
'        End If
        
    End If
    
End Function

Public Property Get PSCReadMeFile() As String
Attribute PSCReadMeFile.VB_Description = "PSC Zip Files Only: returns the name of the PSC Read Me Text File included in the Zip."
    PSCReadMeFile = mPSCFile
End Property

Public Function ReadZip(ByVal pTheZipFileName As String, Optional ByVal pvLoadListView As Boolean = False) As Boolean
Attribute ReadZip.VB_Description = "Interface Instruction to Read a Zip File and process stats etc with option to Load Data into ListView."
Dim xOptions As cOptions
Dim bExtract As Boolean
    
    If Dir$(pTheZipFileName, vbNormal) <> "" Then
        
        ReadZip = pOpen(pTheZipFileName, pvLoadListView)
        
        If ReadZip = True Then
            Set xOptions = New cOptions
            xOptions.Read
            bExtract = xOptions.AutoUnZip
            Set xOptions = Nothing
            If bExtract Then
                cmdExtract_Click
            End If
        End If
    End If

End Function

Public Property Get UnzipFolder() As String
Attribute UnzipFolder.VB_Description = "Returns the name of the Folder where the decompressed Zip File Contents were written."
    UnzipFolder = mUnzipFolder
End Property

Public Property Get VBPFileString() As String
Attribute VBPFileString.VB_Description = "Returns a vbCrLF delimited string with the names of the VBP Files found in the Zip File."
    VBPFileString = mVBPFileString
End Property

Private Sub pWriteUnzipPSCDetails()
' write psc readme details to file
Dim sTmp As String
Dim xString As SBuilder ' StringWorker
Dim xOptions As cOptions
Dim bOK As Boolean
Dim sErrMsg As String

    If Len(mtPSCInfo.Name) = 0 Then GoTo Quit:
    Set xString = New SBuilder ' StringWorker
    Set xOptions = New cOptions
    xOptions.Read bOK, sErrMsg
    If bOK Then
        sTmp = xOptions.UnzipFolder
        xString.ReadFromFile sTmp & "\PSCHistory.txt"
        With mtPSCInfo
            If xString.Length > 0 Then xString.Append Chr$(3)
            xString.Append .Name & Chr$(2) & .UnzipFolder & Chr$(2) & .Link & Chr$(2) & .Description
        End With
        xString.WriteToFile sTmp & "\PSCHistory.txt"
    End If
    Set xOptions = Nothing
    Set xString = Nothing
Quit:

End Sub

Private Sub imgBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' -------------------------------------------------------------------
' ... Toolbar Image: Shift Right and Down.
' -------------------------------------------------------------------
    If mLMouseDown Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left + 15, imgBtn(Index).Top + 15
    mLMouseDown = True
End Sub

Private Sub imgBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' -------------------------------------------------------------------
' ... Toolbar Image: Shift Left and Up.
' -------------------------------------------------------------------
    If mLMouseDown = False Then Exit Sub
    imgBtn(Index).Move imgBtn(Index).Left - 15, imgBtn(Index).Top - 15
    mLMouseDown = False
End Sub

Private Sub pLoadButtonCursors()

' ... sets up the hand cursor for the colour labels.

Dim lngCount As Long
Dim lngLoop As Long

    If myHand_handle = 0 Then Exit Sub
    
    On Error Resume Next ' ... in case the indexes are out.
    
    lngCount = imgBtn.Count
    For lngLoop = 0 To lngCount - 1
        imgBtn(lngLoop).MouseIcon = myHandCursor
        imgBtn(lngLoop).MousePointer = vbCustom
    Next lngLoop
    
End Sub

