VERSION 5.00
Begin VB.Form frmCompileResource 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mini Manifest Resource maker"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   38
   Icon            =   "frmCompileResource.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHookText 
      Alignment       =   1  'Right Justify
      Caption         =   "Include implementation code text file"
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Check to include implementation code in text file."
      Top             =   4200
      Width           =   3315
   End
   Begin VB.CommandButton cmdCreateResource 
      Caption         =   "Create Resource"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Create a Resource file"
      Top             =   4140
      Width           =   1635
   End
   Begin VB.CommandButton cmdSelectVBP 
      Caption         =   "Select VBP"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Select a VB5/6 Project"
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label Label11 
      Caption         =   "2. Check / Uncheck ' Include Implementation code... ' to write hook code to text file."
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   984
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Caption         =   "6. Recompile the program executable."
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "4. [ Copy the Implementation code into the project and hook it up. ]"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1662
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "5. Add the Resource File to the VB5/6 Project."
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2001
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "3. Click ' Generate Resource ' to create the Resource file."
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1323
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Caption         =   "1. Click ' Select VBP ' to select the VB Project file."
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   645
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Create a Resource File with a built-in Manifest to add to a VB5/6 Project."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   6435
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblResPath 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   6435
   End
   Begin VB.Label lblVBPath 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3420
      Width           =   6435
   End
End
Attribute VB_Name = "frmCompileResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A visual interface for creating a manifest compiled into a resource file for a project."

' what?
'  a visual user interface to the Generate Manifest Resource function.
' why?
'  to help create a compiled Manifest within a Resource File
'  with a view to leveraging XP Themes without requiring a Manifest.
' when?
'  there is a project without a Resource File already attached.
' how?
'  load the form, select a vbp and click ' Create Resource '
' who?
'  d.c.

Option Explicit

Dim mVBPPath As String
'Dim mFromProjExplorer As Boolean

Sub SetUpVBP(pvVBPFile As String)
    If modFileName.FileExists(pvVBPFile) Then
        lblVBPath.Caption = pvVBPFile
        mVBPPath = pvVBPFile
'        mFromProjExplorer = True
    End If
End Sub

Private Sub cmdCreateResource_Click()
Attribute cmdCreateResource_Click.VB_Description = "attempts to create the resource file with compiled manifest."

Dim sPath As String
Dim sResPath As String
Dim sResEntry As String

Dim bOK As Boolean
Dim sErrMsg As String
Dim oVBPInfo As VBPInfo
Dim iAnswer As VbMsgBoxResult
Dim xString As SBuilder ' StringWorker
Dim iFind As Long
Dim iFindEnd As Long
Dim sFind As String
Dim iLen As Long

    cmdCreateResource.Enabled = False
    Me.Enabled = False
    
    If Len(mVBPPath) Then
        
        If Dir$(mVBPPath, vbNormal) <> "" Then
            
            modManifestRes.GenerateManifestResource mVBPPath, sPath, chkHookText.Value And vbChecked, bOK, sErrMsg
        
            If bOK = True Then
            
                lblResPath.Caption = sPath
                ' v8, see about adding the res file to the project. -----------------
                Set oVBPInfo = New VBPInfo
                oVBPInfo.ReadVBP mVBPPath
                
                If oVBPInfo.IsExe And oVBPInfo.HasResource = False Then
                    iAnswer = MsgBox("Add this resource file to the Project's VBP?", vbQuestion + vbYesNo, Caption)
                    If iAnswer = vbYes Then
                        ' -------------------------------------------------------------------
                        FileCopy mVBPPath, mVBPPath & ".bkp"
                        sFind = "ResFile32="
                        ' -------------------------------------------------------------------
                        Set xString = New SBuilder 'StringWorker
                        xString.ReadFromFile mVBPPath
                        ' -------------------------------------------------------------------
                        iFind = xString.Find(sFind)
                        If iFind Then
                            sFind = vbNewLine
                            iFindEnd = xString.Find(sFind, iFind + 1)
                            iLen = iFindEnd - iFind
                            If iLen > 0 Then
                                iLen = iLen + 2 ' account for new line chars
                                ' delete existing res ref -------------------------------------------
                                xString.DeletePortion iFind, iLen
                            End If
                        Else
                            sFind = "Form="
                            iFind = xString.Find(sFind)
                        End If
                        ' -------------------------------------------------------------------
                        If iFind > 0 Then
                            ' update the vbp ----------------------------------------------------
                            sResPath = Mid$(sPath, Len(oVBPInfo.FilePath) + 2)
                            sResEntry = "ResFile32=" & Chr$(34) & sResPath & Chr$(34) & vbNewLine
                            ' -------------------------------------------------------------------
                            xString.Insert sResEntry, iFind, bOK, sErrMsg
                            ' -------------------------------------------------------------------
                            If bOK Then
                                xString.WriteToFile mVBPPath, , bOK, sErrMsg
                                If bOK Then
                                    ' -------------------------------------------------------------------
                                    MsgBox oVBPInfo.Title & " has been updated with Resource" & vbNewLine & sPath, vbInformation, Caption
                                Else
NoUpdate:
                                    MsgBox oVBPInfo.Title & " has Not been updated with Resource" & vbNewLine & sErrMsg, vbInformation, Caption
                                End If
                            Else
                                GoTo NoUpdate:
                            End If
                        End If
                        ' -------------------------------------------------------------------
                        Set xString = Nothing
                        ' -------------------------------------------------------------------
                    End If
                End If
                            
                Set oVBPInfo = Nothing
                ' -------------------------------------------------------------------
            Else
                
                MsgBox "The following reason was given for not processing the manifest request:" & vbNewLine & sErrMsg, vbInformation, "Manifest Resource Compile"
                
            End If
        
        End If
        
    End If

    cmdCreateResource.Enabled = True
    Me.Enabled = True

End Sub ' ... cmdCreateResource_Click:

Private Sub cmdSelectVBP_Click()
Attribute cmdSelectVBP_Click.VB_Description = "loads the Open File Name Dialog Box to select a VB Project."

Dim sFileName As String
Dim sCDFilter As String
        
    sCDFilter = modDialog.MakeDialogFilter("Visual Basic Project", , "vbp")

    sFileName = modDialog.GetOpenFileName(, , sCDFilter)
    
    If Len(sFileName) Then

        lblVBPath.Caption = sFileName
        mVBPPath = sFileName
        
    End If
    
End Sub ' ... cmdSelectVBP_Click:

Private Sub Form_Load()
Attribute Form_Load.VB_Description = "sets the current VB Project Path to null string."

    mVBPPath = ""

End Sub ' ... Form_Load:

Private Sub Form_Unload(Cancel As Integer)
Attribute Form_Unload.VB_Description = "releases resources used."

    mVBPPath = ""
    ClearMemory
    
End Sub ' ... Form_Unload:
