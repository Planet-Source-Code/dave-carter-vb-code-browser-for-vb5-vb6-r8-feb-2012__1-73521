VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCopyProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Project"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   39
   Icon            =   "frmCopyProject.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      Height          =   675
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmCopyProject.frx":058A
      ToolTipText     =   "The new project's description"
      Top             =   2760
      Width           =   6615
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Title"
      ToolTipText     =   "The new project's Title"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CheckBox chkVB5 
      Alignment       =   1  'Right Justify
      Caption         =   "Include VB5 VBP?"
      Height          =   375
      Left            =   3540
      TabIndex        =   4
      Top             =   2220
      Width           =   1635
   End
   Begin VB.CheckBox chkZip 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip New Project"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   1680
      Width           =   1635
   End
   Begin ComctlLib.ListView lv 
      Height          =   2235
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Type"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Path"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "Remove File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3900
      TabIndex        =   12
      Top             =   7680
      Width           =   1395
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   7680
      Width           =   1395
   End
   Begin VB.CommandButton cmdOpenCopy 
      Caption         =   "Open Folder"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdCopyProj 
      Caption         =   "Copy Project"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   1620
      Width           =   1395
   End
   Begin VB.CommandButton cmdSelProject 
      Caption         =   "Select Project"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1620
      Width           =   1395
   End
   Begin VB.Label lblProgress 
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInstructions 
      Caption         =   "3. Click ' Open Folder '' to open the copied project's folder."
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   10860
      Visible         =   0   'False
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInstructions 
      Caption         =   "2. Click ' Copy Project ' to begin the copy process."
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   10500
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInstructions 
      Caption         =   "1. Select a Project (Click ' Select Project ' to choose VBP from Open File Dialog )."
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   10140
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmCopyProject.frx":0596
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProjSource 
      Caption         =   "Path to Source Project:"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProjSource 
      Caption         =   "Source Project:"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   6615
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCopyProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for copying a project to a new folder or zip."

' what?
'  user interface to copying a project.
'  the intention is that only the required files
'  exist in the new project's directory.
' why?
'  developing a project can lead to extraneous files
'  that were used along the way but are no longer in use.
'  when zipping the project to pass it on, one might prefer
'  that any extraneous files are not included in order to minimise
'  the size of the project.
' when?
'  the developer wishes to reduce the size of a project's
'  footprint on disk, perhaps for zipping to PSC or maybe
'  the project's folders have become too cluttered.
' how?
'  select a project and click the Copy Project button.
' who?
'  dc.

' Note:
'  the new project will be written to the source project's
'  parent folder so that references and objects can be
'  copied directly to the new vbp.
'  the new project's folders will include a folder each
'  for the forms, classes, modules and user controls of the source project.
'  Property Pages, Data Environment and Data Reports may not make
'  it into the process until a later date.

' Update: 09 April 2011
'  v5 was incomplete in its range of files that it copied.
'  this was busker's version 1, it remains, being updated to
'  include some missing things such as property pages.
'  v6 offers option to zip the newly created project
'  and add files to the copy process.
'  Busker's version 2, v6, has progressed toward being
'  a bit more complete; reading binary files is important
'  to providing something more solid and true.
'  This version is still in long hand in the meantime.

'  The copy routine does not allow renaming of the project name
'  only Title and Description.
'  If we change the name we should then check all forms and user controls
'  for use of user controls within the project and, if found, then
'  change the reference to the control within the source header (replace old project name with new).

Option Explicit

Private Const cDlgCancelErr As Long = 32755

Private mLines() As Long
Private mText As String
Private mFileNameInfo As FileNameInfo
Private moVBPInfo As VBPInfo
Private mInitialised As Boolean
Private mUnloadPostCopy As Boolean
Private mLastAddFolder As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub LoadVBP(pTheVBPFileName As String)
    
    psOpen pTheVBPFileName
    
    ' ... wanted copy button to have focus
    ' ... kept restoring to select proj button
    cmdSelProject.Enabled = False
    mUnloadPostCopy = True
'    On Error Resume Next
'    cmdCopyProj.SetFocus
    
End Sub

Private Sub pRelease()
    
    If Not moVBPInfo Is Nothing Then
        Set moVBPInfo = Nothing
    End If
    
    If Not lv.SmallIcons Is Nothing Then
        Set lv.SmallIcons = Nothing
    End If
    
    mInitialised = False
    mText = vbNullString
    mUnloadPostCopy = False
        
    mLastAddFolder = vbNullString
    
    Erase mLines
    
End Sub

Private Sub pInit()
    
    pRelease
    
    lblProjSource(0).Caption = ""
    lblProjSource(1).Caption = ""
    
    lblProgress.Caption = ""
    
    txtTitle.Text = ""
    txtDescription.Text = ""
    
End Sub

Private Sub psOpen(Optional pTheFileName As String = vbNullString)

Dim sFileName As String
Dim sCDFilter As String
Dim sDel As String
Dim sProjInfo As String
Dim bOK As Boolean
Dim bIsVBP As Boolean

    On Error GoTo ErrHan:
    
    pInit
    
    cmdCopyProj.Enabled = False
    cmdOpenCopy.Enabled = False
    cmdSelProject.Enabled = False
    cmdAddFile.Enabled = False
    cmdRemoveFile.Enabled = False
    chkZip.Enabled = False
    
    If pTheFileName = vbNullString Then
    
        sCDFilter = modDialog.MakeDialogFilter("Visual Basic Project", , "vbp")
    
        sFileName = modDialog.GetOpenFileName(, , sCDFilter)
    
    Else
        
        sFileName = pTheFileName
    
    End If
    
    If Len(sFileName) Then
                
        modReader.ReadTextFile sFileName, mText, mLines, sDel
        
        modFileName.ParseFileNameEx sFileName, mFileNameInfo
        
        With mFileNameInfo
        
            If LCase$(.Extension) = "vbp" Then
                
                bIsVBP = True
                
                Set moVBPInfo = New VBPInfo
                moVBPInfo.ReadVBP .PathAndName
                
                bOK = moVBPInfo.Initialised
                
                If bOK = True Then
                
                    sProjInfo = "Source Project:" & " " & moVBPInfo.ProjectName & ", " & moVBPInfo.FileName & vbNewLine
                    sProjInfo = sProjInfo & "Project Type:" & "    " & moVBPInfo.ProjectType & vbNewLine
'                    sProjInfo = sProjInfo & "Project Desc:" & "    " & moVBPInfo.Description
                    
                    lblProjSource(0).Caption = sProjInfo
                    lblProjSource(1).Caption = "Source Path:" & "    " & moVBPInfo.FileNameAndPath
                    
                    txtTitle.Text = moVBPInfo.Title
                    txtDescription.Text = moVBPInfo.Description
                    
                End If
                
            End If
            
        End With
    
    End If
        
ResErr:
    
    mInitialised = bOK
    
    If bOK = False Then
    
        If Not moVBPInfo Is Nothing Then
            
            Set moVBPInfo = Nothing
            
        End If
        
        If bIsVBP Then
            MsgBox "The VB Project file wasn't opened properly" & vbNewLine & "and not able to continue.", vbExclamation, Caption
        End If
    
    Else
    
        cmdCopyProj.Enabled = True
        cmdAddFile.Enabled = True
        ' -------------------------------------------------------------------
        ' ... see if the Zip32.dll file is available in sys32.
        If modGeneral.CheckZipAPI Then
            chkZip.Enabled = True
        Else
            chkZip.Value = vbUnchecked
        End If
        
    End If
    
    ' ... even if open was unsuccessful, if there are files
    ' ... in the listview then let them be removed.
    cmdRemoveFile.Enabled = CBool(lv.ListItems.Count > 0)
    
    cmdSelProject.Enabled = True
    
    sProjInfo = vbNullString
    sCDFilter = vbNullString
    sFileName = vbNullString
    
Exit Sub
ErrHan:
    bOK = False
    If Err.Number <> cDlgCancelErr Then
        Debug.Print Err.Number
    End If
    Resume ResErr:
    
End Sub

Private Function pbCopyFile(pTheSourceFile As String, pTheDestinationFile As String, Optional pErrMsg As String = vbNullString) As Boolean

' ... helper to copy a file, returns false if unsuccessful.

    On Error GoTo ErrHan:
    
    If Len(pTheSourceFile) = 0 Then Err.Raise vbObjectError + 1000, , "Source File Name missing."
    If Len(pTheDestinationFile) = 0 Then Err.Raise vbObjectError + 1000, , "Destination File Name missing."
    If FileLen(pTheSourceFile) = 0 Then Err.Raise vbObjectError + 1000, , "Source File is empty or non-existent."
    
    FileCopy pTheSourceFile, pTheDestinationFile
    
    pbCopyFile = True

Exit Function

ErrHan:

    pErrMsg = Err.Description
    Debug.Print "frmCopyProject.pbCopyFile.Error: " & Err.Number & "; " & Err.Description

End Function

Private Sub cmdAddFile_Click()

Dim sFile As String
Dim xFileInfo As FileNameInfo
Dim xItem As ListItem
Dim sItem As String
Dim sFileInfo As String
Dim sInitDir As String

    On Error GoTo ErrHan:
    ' -------------------------------------------------------------------
    ' ... if not called yet, check if we have a project path
    ' ... and then set this to the initial directory for easier
    ' ... navigation to required files.
    
    If Len(mLastAddFolder) Then
        sInitDir = mLastAddFolder
    Else
        If mFileNameInfo.Path <> "" Then
            sInitDir = mFileNameInfo.Path
        End If
    End If
    
    ' -------------------------------------------------------------------
    sFile = modDialog.GetOpenFileName(, , , , sInitDir, "Add File")
    
    If Len(sFile) Then
        If lv.SmallIcons Is Nothing Then
            Set lv.SmallIcons = mdiMain.ilProject
        End If
        modFileName.ParseFileNameEx sFile, xFileInfo
        If Len(xFileInfo.Path) Then
            mLastAddFolder = xFileInfo.Path
        End If
        sFileInfo = vbNullString
        sItem = modFileIcons.AddIconToImageList(sFile, mdiMain.ilProject, xFileInfo.Extension, sFileInfo)
        Set xItem = lv.ListItems.Add(, xFileInfo.PathAndName, xFileInfo.File, , sItem)
        xItem.SubItems(1) = sFileInfo
        xItem.SubItems(2) = xFileInfo.Path
        
    End If

ResumeError:
    
    cmdRemoveFile.Enabled = lv.ListItems.Count > 0
    
    If Not xItem Is Nothing Then
        Set xItem = Nothing
    End If
    
Exit Sub

ErrHan:
    If Err.Number <> 35602 Then ' ... key not unique, file already added.
        Debug.Print "frmCopyProject.cmdAddFile_Click.Error: " & Err.Number & "; " & Err.Description
    End If
    Resume ResumeError:

End Sub

Private Sub cmdCopyProj_Click()

' ... Copy Project, Busker's version.
' ... essentially includes most stuff to perform the functiom
' ... caution, crap coding; i'll sort it out later.

Dim lngCount As Long
Dim lngLoop As Long
Dim xFileNameInfo As FileNameInfo
Dim soTmpArray As StringArray
Dim sTmp As String

Dim xString As SBuilder ' StringWorker
Dim x5String As SBuilder ' StringWorker

Dim sVBPFileLines As String
Dim sVB5VBPFileLines As String

Dim sName As String
Dim sTargetPath As String
Dim iAnswer As VbMsgBoxResult
Dim sTmpFolder As String
Dim bOK As Boolean
Dim sErrMsg As String
Dim sTmpExt As String
Dim bEnabled As Boolean
Dim xItem As ListItem
Dim bHasPropPages As Boolean
Dim bHasUCs As Boolean
Dim bIncVB5 As Boolean
Dim bDidVB5 As Boolean
Dim sTmpFile As String

Const cBinFolder As String = "bcp"

' -------------------------------------------------------------------
' Note: Original Busker's version, 1 (scoping), updated for v6,
'       now attempts a zip on the new folder and includes
'       missing file types from before.
'       K, so fair amount of scoping achieved, if all
'       turns out fine then can rewrite this function.
' -------------------------------------------------------------------

    On Error GoTo ErrHan:
    
    cmdCopyProj.Enabled = False
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    If mInitialised = False Then Err.Raise vbObjectError + 1000, , "Not Initialised."
    If moVBPInfo Is Nothing Then Err.Raise vbObjectError + 1000, , "VBPInfo Not Instanced."
    If moVBPInfo.Initialised = False Then Err.Raise vbObjectError + 1000, , "VBPInfo Not Initialised."

    lngCount = moVBPInfo.CountVBPFiles
    If lngCount < 1 Then Err.Raise vbObjectError + 1000, , "No File References in VBP."
    
    sTargetPath = moVBPInfo.FilePath & " Copy"
    lblProgress = "Checking New Project Folders"
    lblProgress.Refresh
    
    If Dir$(sTargetPath, vbDirectory) <> "" Then
        iAnswer = MsgBox("The Target Folder " & sTargetPath & " already Exists," & vbNewLine & "OK to overwrite files?", vbQuestion + vbYesNo, Caption)
        If iAnswer = vbNo Then
            Err.Raise vbObjectError + 1000, , "Target Folder Exists and the operation was cancelled."
        End If
    Else
        MkDir sTargetPath
    End If
    
    bIncVB5 = chkVB5.Value And vbChecked
    ' -------------------------------------------------------------------
    ' ... long hand directory checking.
    If Dir$(sTargetPath & "\frm", vbDirectory) = "" Then
        MkDir sTargetPath & "\frm"
    End If
    If Dir$(sTargetPath & "\cls", vbDirectory) = "" Then
        MkDir sTargetPath & "\cls"
    End If
    If bIncVB5 Then
        If Dir$(sTargetPath & "\cls\vb5", vbDirectory) = "" Then
            MkDir sTargetPath & "\cls\vb5"
        End If
    End If
    If Dir$(sTargetPath & "\mod", vbDirectory) = "" Then
        MkDir sTargetPath & "\mod"
    End If
    If Dir$(sTargetPath & "\uc", vbDirectory) = "" Then
        MkDir sTargetPath & "\uc"
    End If
    ' -------------------------------------------------------------------
'    ' v6, property pages
'    If Dir$(sTargetPath & "\ppg", vbDirectory) = "" Then
'        MkDir sTargetPath & "\ppg"
'    End If
    ' ... data environment / report (designer).
    If Dir$(sTargetPath & "\des", vbDirectory) = "" Then
        MkDir sTargetPath & "\des"
    End If
    ' -------------------------------------------------------------------
    If Dir$(sTargetPath & "\hlp", vbDirectory) = "" Then
        MkDir sTargetPath & "\hlp"
    End If
    If Dir$(sTargetPath & "\res", vbDirectory) = "" Then
        MkDir sTargetPath & "\res"
    End If
    
    ' -------------------------------------------------------------------
    lblProgress = "Beginning File Copying..."
    lblProgress.Refresh
    ' -------------------------------------------------------------------
    ' ... copy files referenced in vbp.
    
    For lngLoop = 1 To lngCount ' ... lngCount is the number of files to read as given by moVBPInfo
        
        Set soTmpArray = moVBPInfo.FilesData.ItemAsStringArray(lngLoop, "|")
        modFileName.ParseFileNameEx soTmpArray(soTmpArray.Count), xFileNameInfo
        
        sTmp = ""
        
        sName = soTmpArray(1)           ' ... soTmpArray(1) is source file name.
        
        Select Case Val(soTmpArray(2))  ' ... soTmpArray(2) is source file type, see VBPInfo for definitions.
        
            Case 4: sTmp = "Class=" & sName & "; cls\" & xFileNameInfo.File: sTmpFolder = "\cls\"
            Case 3: sTmp = "Module=" & sName & "; mod\" & xFileNameInfo.File: sTmpFolder = "\mod\"
            Case 5: sTmp = "UserControl=uc\" & xFileNameInfo.File: sTmpFolder = "\uc\"
                bHasUCs = True
            Case 21, 22, 23: sTmp = "Form=frm\" & xFileNameInfo.File: sTmpFolder = "\frm\"
            
            Case 10: sTmp = "PropertyPage=uc\" & xFileNameInfo.File: sTmpFolder = "\uc\" ' v6, not linked to user controls after copy process!
'            Case 10: sTmp = "PropertyPage=ppg\" & xFileNameInfo.File: sTmpFolder = "\ppg\" ' v6, not linked to user controls after copy process!
            ' ... copy property pages to user control folder.
            ' ... until update later, e.g. check out prop pages from user control binary files.
                bHasPropPages = True
            Case 7, 8
                sTmp = "Designer=des\" & xFileNameInfo.File: sTmpFolder = "\des\"
                
        End Select
        
        sTmpExt = UCase$(xFileNameInfo.Extension)
        
        If Len(sTmp) Then
            If Len(sVBPFileLines) Then sVBPFileLines = sVBPFileLines & vbNewLine
            sVBPFileLines = sVBPFileLines & sTmp
            If sTmpExt <> "CLS" Then
                If Len(sVB5VBPFileLines) > 0 Then sVB5VBPFileLines = sVB5VBPFileLines & vbNewLine
                sVB5VBPFileLines = sVB5VBPFileLines & sTmp
            End If
        End If
        
        lblProgress = "Copying ..." & vbNewLine & xFileNameInfo.PathAndName
        lblProgress.Refresh
        ' -------------------------------------------------------------------
        ' ... copy the file.
        
        bOK = pbCopyFile(xFileNameInfo.PathAndName, sTargetPath & sTmpFolder & xFileNameInfo.File, sErrMsg)
        
        ' -------------------------------------------------------------------
        
        ' ... look out for binary files to include.
        
        If sTmpExt = "FRM" Or sTmpExt = "CTL" Or sTmpExt = "PAG" Or sTmpExt = "DSR" Then
            
            If sTmpExt = "FRM" Then
                sTmpExt = ".frx"
            
            ElseIf sTmpExt = "CTL" Then
                sTmpExt = ".ctx"
            
            ElseIf sTmpExt = "PAG" Then ' v6, pick up property pages and their binary data.
                sTmpExt = ".pgx"
            
            ElseIf sTmpExt = "DSR" Then
                sTmpExt = ".dsx"
                
            Else
                sTmpExt = vbNullString
            End If
            
ResumeWithBinaryFile:

            sTmp = xFileNameInfo.Path & "\" & xFileNameInfo.FileName & sTmpExt
            
            If Dir$(sTmp, vbNormal) <> "" And Len(sTmpExt) Then
        
                lblProgress = "Copying ..." & vbNewLine & sTmp
                lblProgress.Refresh
                ' -------------------------------------------------------------------
                ' ... copy the binary file companion.
                bOK = pbCopyFile(sTmp, sTargetPath & sTmpFolder & xFileNameInfo.FileName & sTmpExt, sErrMsg)
                
                If sTmpExt = ".dsx" Then
                    ' ... data environment / report, look up dca file.
                    sTmpExt = ".dca"
                    GoTo ResumeWithBinaryFile:
                    
                End If
                
            End If
        
        ElseIf sTmpExt = "CLS" And bIncVB5 Then
            ' -------------------------------------------------------------------
            ' ... note, some class header attributes in vb6 are not valid in vb5.
            ' ... create a duplicate to act as surrogate class in vb5 project.
            sTmpFolder = "\cls\vb5\"
            sTmp = "Class=" & sName & "; cls\vb5\" & xFileNameInfo.File
            
            If Len(sVB5VBPFileLines) > 0 Then sVB5VBPFileLines = sVB5VBPFileLines & vbNewLine
            sVB5VBPFileLines = sVB5VBPFileLines & sTmp
            
            sTmp = sTargetPath & sTmpFolder & xFileNameInfo.File
            
            bOK = pbCopyFile(xFileNameInfo.PathAndName, sTmp, sErrMsg)
            
            Set x5String = New SBuilder ' StringWorker
            
            x5String.ReadFromFile sTmp
            ' -------------------------------------------------------------------
            ' ... well dodgey bit of logic, potentially harmfu to source in class
            ' ... if these exact values are duplicated elsewhere.
            ' ... Replace should like to provide a region and number of replacements.
            x5String.Replace "  Persistable = 0  'NotPersistable" & vbCrLf, ""
            x5String.Replace "  DataBindingBehavior = 0  'vbNone" & vbCrLf, ""
            x5String.Replace "  DataSourceBehavior  = 0  'vbNone" & vbCrLf, ""
            x5String.Replace "  MTSTransactionMode  = 0  'NotAnMTSObject" & vbCrLf, ""
            x5String.WriteToFile sTmp
            
            Set x5String = Nothing
            
            bDidVB5 = True
            
        End If
    
    Next lngLoop
    
    ' -------------------------------------------------------------------
    ' ... Related Docs, v6
    Set soTmpArray = moVBPInfo.RelatedDocs
    If soTmpArray.Count > 0 Then
        For lngLoop = 1 To soTmpArray.Count
            sTmp = soTmpArray(lngLoop)
            If Dir$(sTmp) <> "" Then
                modFileName.ParseFileNameEx sTmp, xFileNameInfo
                
                lblProgress = "Copying ..." & vbNewLine & xFileNameInfo.PathAndName
                lblProgress.Refresh
                sTmpFolder = "\"
                bOK = pbCopyFile(xFileNameInfo.PathAndName, sTargetPath & sTmpFolder & xFileNameInfo.File, sErrMsg)
                
            End If
        Next lngLoop
    End If
    
    lblProgress = "Copying VBP."
    lblProgress.Refresh

    ' -------------------------------------------------------------------
    ' ... copy the other vbp entries.
    Set soTmpArray = moVBPInfo.VBPTextLines
    lngCount = soTmpArray.Count
    
    Set xString = New SBuilder ' StringWorker
    
    xString.AppendAsLine soTmpArray(1)  ' ... first line > Type=
    
    For lngLoop = 2 To lngCount
    
        sTmp = soTmpArray(lngLoop)
        
        If sTmp = "" Then Exit For
        
        If Left$(sTmp, 6) = "Object" Then
        
        ElseIf Left$(sTmp, 9) = "Reference" Then
        
        Else
            If Left$(sTmp, 4) = "Form" Then
                sTmp = ""
            ElseIf Left$(sTmp, 5) = "Class" Then
                sTmp = ""
            ElseIf Left$(sTmp, 6) = "Module" Then
                sTmp = ""
            ElseIf Left$(sTmp, 11) = "UserControl" Then
                sTmp = ""
            ElseIf Left$(sTmp, 12) = "PropertyPage" Then ' v6, added.
                sTmp = ""
            ElseIf Left$(sTmp, 8) = "Designer" Then ' v6, added.
                sTmp = ""
            ElseIf Left$(sTmp, 5) = "Title" Then ' v6, added.
                ' ... provide the new user defined Title.
                sTmp = txtTitle.Text
                
                modStrings.RemoveQuotes sTmp
                
                sTmp = modStrings.WrapInQuoteChars(sTmp)
                
                sTmp = "Title=" & sTmp ' & Chr$(34) & txtTitle.Text & Chr$(34)
                
            ElseIf Left$(sTmp, 11) = "Description" Then ' v6, added.
                ' ... provide the new user defined Description.
                sTmp = txtDescription.Text
                
                modStrings.RemoveQuotes sTmp
                
                sTmp = modStrings.Replace(sTmp, vbCrLf, "")
                sTmp = modStrings.Replace(sTmp, vbLf, "")
                sTmp = modStrings.Replace(sTmp, vbCr, "")
                sTmp = modStrings.WrapInQuoteChars(sTmp)
                
                sTmp = "Description=" & sTmp
                
            ElseIf Left$(sTmp, 9) = "ResFile32" Then
                sTmp = ""
                sTmp = moVBPInfo.ResFileNameAndPath
                If Len(sTmp) Then
                    modFileName.ParseFileNameEx sTmp, xFileNameInfo
                    sTmp = "ResFile32=" & WrapInQuoteChars("res\" & xFileNameInfo.File)
                    sTmpFolder = "\res\"
        
                    lblProgress = "Copying ..." & vbNewLine & xFileNameInfo.PathAndName
                    lblProgress.Refresh
                    
                    bOK = pbCopyFile(xFileNameInfo.PathAndName, sTargetPath & sTmpFolder & xFileNameInfo.File, sErrMsg)
                    
                End If
            ElseIf Left$(sTmp, 8) = "HelpFile" Then
                sTmp = ""
                If Len(moVBPInfo.HelpFile) Then
                    sTmp = moVBPInfo.HelpFile
                    If Dir$(sTmp, vbNormal) <> "" Then
                        modFileName.ParseFileNameEx sTmp, xFileNameInfo
                        sTmp = "HelpFile=" & WrapInQuoteChars(sTargetPath & "\hlp\" & xFileNameInfo.File)
                        sTmpFolder = "\hlp\"
            
                        lblProgress = "Copying ..." & vbNewLine & xFileNameInfo.PathAndName
                        lblProgress.Refresh
                        
                        bOK = pbCopyFile(xFileNameInfo.PathAndName, sTargetPath & sTmpFolder & xFileNameInfo.File, sErrMsg)
                        
                    End If
                End If
            ElseIf Left$(sTmp, 8) = "Retained" Then
                ' ... Note:
                ' ... Ignoring Retained because this invalidates a vb5 vbp,
                ' ... doing so with no understanding of what this does for vb6.
                sTmp = ""
            
            ElseIf Left$(sTmp, 15) = "CompatibleEXE32" Then
            
                ' ... ensure VersionCompatible32 = 1 before continuing.
                ' ... if not in use then no need to copy binary compat version.
                
                If moVBPInfo.VersionCompatible32 = "1" Then
                    
                    modStrings.SplitStringPair sTmp, "=", sTmp, sTmpExt
                    modStrings.RemoveQuotes sTmpExt
                    modVB.ReadVBFilePath moVBPInfo.FilePath, sTmpExt
                    
                    If Dir$(sTargetPath & "\" & cBinFolder, vbDirectory) = "" Then
                        MkDir sTargetPath & "\" & cBinFolder ' ... bcp: binary compatibility
                    End If
                    
                    If Dir$(sTargetPath & "\" & cBinFolder, vbDirectory) <> "" Then
                        
                        modFileName.ParseFileNameEx sTmpExt, xFileNameInfo
                        
                        lblProgress = "Copying ..." & vbNewLine & xFileNameInfo.PathAndName
                        lblProgress.Refresh
                        ' -------------------------------------------------------------------
                        pbCopyFile sTmpExt, sTargetPath & "\" & cBinFolder & "\" & xFileNameInfo.File
                        
                        ' -------------------------------------------------------------------
                        ' ... check out supporting files.
                        sTmpFile = Dir$(xFileNameInfo.Path & "\" & xFileNameInfo.FileName & ".*", vbNormal)
                        
                        Do While Len(sTmpFile)
                            
                            If LCase$(xFileNameInfo.Path & "\" & sTmpFile) <> LCase$(sTmpExt) Then
                            
'                                Debug.Print xFileNameInfo.Path & "\" & sTmpFile & " to " & sTargetPath & "\" & cBinFolder & "\" & sTmpFile
                                ' -------------------------------------------------------------------
                                pbCopyFile xFileNameInfo.Path & "\" & sTmpFile, sTargetPath & "\" & cBinFolder & "\" & sTmpFile
                            
                            End If
                            
                            sTmpFile = Dir$
                        
                        Loop
                    
                    End If
                    
                    ' ... vbp item.
                    sTmp = sTmp & "=" & modStrings.WrapInQuoteChars(cBinFolder & "\" & xFileNameInfo.File)
                
                End If
            
            End If
        
        End If
        
        If Len(sTmp) Then
            xString.AppendAsLine sTmp
        End If
        
    Next lngLoop
    
    If Not xString Is Nothing Then
        
        If Len(sVBPFileLines) Then
        
            lblProgress = "Completing VBP"
            lblProgress.Refresh
            
            Set x5String = New SBuilder ' StringWorker
            x5String.TheString = xString
            
            xString.AppendAsLine sVBPFileLines
            x5String.AppendAsLine sVB5VBPFileLines
            
        End If
        
        
        lblProgress = "Writing VBP... " & vbNewLine & sTargetPath & "\" & moVBPInfo.FileName
        lblProgress.Refresh
        
        xString.WriteToFile sTargetPath & "\" & moVBPInfo.FileName, True
        If bIncVB5 Then
            x5String.WriteToFile sTargetPath & "\vb5" & moVBPInfo.FileName, True
        End If
        
    End If
    
    cmdOpenCopy.Tag = sTargetPath
    
    lblProgress = "Project Copied to:" & vbNewLine & sTargetPath
    lblProgress.Refresh
    ' -------------------------------------------------------------------
    ' ... PSC ReadMe file located?
    sTmp = Dir$(moVBPInfo.FilePath & "\@PSC_Read*.txt", vbNormal)
    If sTmp <> "" Then
        
        lblProgress = "Copying PSC ReadMe file:" & vbNewLine & moVBPInfo.FilePath & "\" & sTmp ' sTargetPath & "\" & moVBPInfo.Filename
        lblProgress.Refresh
        
        bOK = pbCopyFile(moVBPInfo.FilePath & "\" & sTmp, sTargetPath & "\" & sTmp, sErrMsg)
    
    End If
    ' -------------------------------------------------------------------
    ' ... Added Files, v6.
    If lv.ListItems.Count Then
        lblProgress = "Copying added files to " & sTargetPath & vbNewLine & moVBPInfo.Title
        lblProgress.Refresh
        For Each xItem In lv.ListItems
            sTmp = xItem.Key
            If Len(sTmp) Then
                If Dir$(sTmp) <> "" Then
                    
                    modFileName.ParseFileNameEx sTmp, xFileNameInfo
                    
                    lblProgress = "Checking... " & xFileNameInfo.File & vbNewLine & "to " & sTargetPath
                    lblProgress.Refresh
                    
                    ' ... if the file already exists (it shouldn't) then ignore its copy.
                    ' ... this is a safeguard against overwriting one of the newly copied files
                    ' ... without having to mess around with validation each time a file is added.
                    
                    If Dir$(sTargetPath & "\" & xFileNameInfo.File) = "" Then
                        
                        lblProgress = "Copying... " & xFileNameInfo.File & vbNewLine & "to " & sTargetPath
                        lblProgress.Refresh
                    
                        bOK = pbCopyFile(xFileNameInfo.PathAndName, sTargetPath & "\" & xFileNameInfo.File, sErrMsg)
                    
                    End If
                End If
            End If
        Next xItem
    
    End If
    
    ' -------------------------------------------------------------------
    DoEvents    ' ... play catch up on all that file writing.
    ' -------------------------------------------------------------------
    ' Zip, v6.
    If chkZip And vbChecked Then
        If modGeneral.CheckZipAPI Then
            
            lblProgress = "Project Copy Complete, generating Zip:" & vbNewLine & moVBPInfo.Title
            lblProgress.Refresh
        
            ' ... attempt to zip the new folder.
            modZip.VBZipEx sTargetPath & "\" & moVBPInfo.ProjectName & ".zip" ', pWithDebug = True for debug info.
            ' -------------------------------------------------------------------
            DoEvents    ' ... play catch up again.
            ' -------------------------------------------------------------------
            
        End If
    End If
    
    lblProgress = "Project Copy Complete:" & vbNewLine & moVBPInfo.Title
    lblProgress.Refresh
    
    ' -------------------------------------------------------------------
    If Len(cmdOpenCopy.Tag) Then
        bEnabled = Dir$(cmdOpenCopy.Tag, vbDirectory) <> ""
        cmdOpenCopy.Enabled = bEnabled
        If bEnabled Then
            bOK = True
            cmdOpenCopy_Click
        End If
    End If
    
ResumeError:
    
    If Not xString Is Nothing Then
        Set xString = Nothing
    End If
    
    If Not x5String Is Nothing Then
        Set x5String = Nothing
    End If
    
    cmdCopyProj.Enabled = True
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
    If bOK And mUnloadPostCopy Then
    
        ' ... loads of message boxes.
        If bDidVB5 Then
            MsgBox "Project Copy generated VB5 Classes and you may like to check these are OK.", vbInfoBackground, Caption
        End If
        
        If bHasPropPages And bHasUCs Then
            MsgBox "Project Copy included Property Pages:" & vbNewLine & "Check User Controls." & vbNewLine & "The Copy Method may have broken the link between User Controls and their Property Pages.", vbInformation, Caption
        End If
        
        MsgBox "Project Copied:" & vbNewLine & moVBPInfo.Title & vbNewLine & "to: " & sTargetPath, vbInformation, Caption
        
        Unload Me
        
    End If
    
    ClearMemory
    
Exit Sub

ErrHan:

    MsgBox "An Error was trapped copying the project;" & vbNewLine & Err.Description, vbExclamation, Caption
    
    Debug.Print "frmCopyProject.cmdCopyProj_Click.Error: " & Err.Number & "; " & Err.Description
    Resume ResumeError:

End Sub

Private Sub cmdOpenCopy_Click()

Dim lngRet As Long
Dim iAnswer As VbMsgBoxResult
    
    On Error Resume Next
    
    If Len(cmdOpenCopy.Tag) Then
                
        iAnswer = MsgBox("Open the new Project's Folder?", vbQuestion + vbYesNo, Caption)
        
        If iAnswer = vbYes Then
            lngRet = ShellExecute(0&, vbNullString, cmdOpenCopy.Tag, vbNullString, vbNullString, vbNormalFocus)
        End If
        
    End If
    
    lngRet = 0&
    iAnswer = 0&
    
End Sub

Private Sub cmdRemoveFile_Click()
    
    If Not lv.SelectedItem Is Nothing Then
        lv.ListItems.Remove lv.SelectedItem.Index
        cmdRemoveFile.Enabled = lv.ListItems.Count > 0
    End If
    
End Sub

Private Sub cmdSelProject_Click()
    psOpen
End Sub

Private Sub Form_Load()

    Set Icon = mdiMain.Icon
    Set lv.SmallIcons = mdiMain.ilProject
    
    LVFullRowSelect lv.hwnd
    ClearMemory
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pRelease
    ClearMemory
End Sub

