Attribute VB_Name = "modRde"
Attribute VB_Description = "Code written by Rde on PSC, Rohan Edwards, thank you."

' ... Written and submitted to PSC by Rde
' ... Too well presented and written to do anything but use as is.
' ... http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72521&lngWId=1
' ... Thank you, again, Rohan :)

Option Explicit

Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long

Private Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long

Private Declare Function fGetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const DIR_SEP As String = "\"

Private Const INVALID_FILE_ATTRIBUTES = (-1)

  
' ... I plan to make a little update to ensure that the path is not the drive root folder
' ... e.g. a:\, b:\, c:\, d:\ e.t.c.
' ... also guard against trying to delete windows and sys32 or syswow(whatever).
' ... I have no idea how to determine the name of the program files folder.

 '-----------------------------------------------------

  


 ' This is a Kill Folder function with persistence.

  

 ' It will remove all sub-folders and files and then

 ' optionally delete the specified folder.

  

 ' I found when removing all the files in the temp folder

 ' that some locked files would fail and cause it to not

 ' continue with the rest of the files.

  

 ' This function will continue to remove all unlocked files,

 ' even after finding locked files. However, if locked files

 ' are found, the parent folder will also not get removed.

  


 '-----------------------------------------------------


  

Public Function AddBackslash(sPath As String) As String
Attribute AddBackslash.VB_Description = "Returns a file path with a trailing backslash."

   If Right$(sPath, 1&) = DIR_SEP Then

      AddBackslash = sPath

   Else

      AddBackslash = sPath & DIR_SEP

   End If

End Function

  

 '-----------------------------------------------------

  

Public Function FolderExists(sPath As String) As Boolean
Attribute FolderExists.VB_Description = "Tests for the existence of a folder/directory."

   Dim Attribs As Long

   Attribs = GetAttributes(sPath)

   If Not (Attribs = INVALID_FILE_ATTRIBUTES) Then

      FolderExists = ((Attribs And vbDirectory) = vbDirectory)

   End If

End Function


' -------------------------------------------------------------------


Private Function psGetWindowsDirectory() As String
Attribute psGetWindowsDirectory.VB_Description = "Private call to GetWindowsDiretory to keep this module independent."

Dim sTmpBuffer As String
Dim lngRetLen As Long

    sTmpBuffer = String$(255, 0)
    lngRetLen = fGetWindowsDirectory(sTmpBuffer, 255)
    If lngRetLen > 0 Then
        sTmpBuffer = Left$(sTmpBuffer, lngRetLen)
        psGetWindowsDirectory = sTmpBuffer
    End If
    
    sTmpBuffer = vbNullString
    lngRetLen = 0&
    
End Function


' -------------------------------------------------------------------


Private Function pbSanityCheck(pPath As String) As Boolean
Attribute pbSanityCheck.VB_Description = "Attempts to determine if a folder is ok to delete e.g. not source code folder, not a windows folders, not a drive..."
' ... Attempts to determine if a folder is ok to delete e.g. not source code folder, not a windows folders, not a drive...

Dim sWinDir As String
Dim sPath As String

' ... adding this to help prevent dickheads like me deleting
' ... their source folder accidentally.
' ... and then considered drive and windows folders.
    
' -------------------------------------------------------------------
' ... Note: this should be called post FolderExists which will
' ...       validate there is a valid folder to work on.
' -------------------------------------------------------------------
    
    ' -------------------------------------------------------------------
    ' ... main things that I is checking up on.
    ' -------------------------------------------------------------------
    ' ... is folder a root drive, such as c:\
    ' ... is folder program/application folder e.g. App.Path
    '     .. yes I did KillFolder on my app.path with an exe, yes it worked
    '     .. yes, all the source was deleted
    '     .. yes, I did have a back-up :D
    ' ... is folder a windows folder, such as c:\windows, c:\windows\system32 ....
    ' -------------------------------------------------------------------
     ' ... is it all a bit too cheap?
     ' -------------------------------------------------------------------
     
    sPath = LCase$(pPath)
    If sPath = LCase$(App.Path) Then GoTo Quit
    
    sWinDir = LCase$(psGetWindowsDirectory)
    If Left$(sPath, Len(sWinDir)) = sWinDir Then GoTo Quit

    If Len(sPath) < 4 Then GoTo Quit
    
    ' -------------------------------------------------------------------
    ' ... if we got here, my crappy checking passed, so return true, sanity check passed.
    pbSanityCheck = True
    
Quit:
    sPath = vbNullString
    sWinDir = vbNullString
    
End Function



 '-----------------------------------------------------

  

Public Function KillFolder(sSpec As String, Optional ByVal bJustEmptyDontRemove As Boolean) As Boolean
Attribute KillFolder.VB_Description = "Attempts to delete a folder, all its files and sub-folders."

   Dim sRoot As String, sDir As String, sFile As String

   Dim iCnt As Long, iIdx As Long

  
'   MsgBox "Make sure its not a drive or system folder", vbInformation, "Kill Folder"
   
   If Not FolderExists(sSpec) Then Exit Function

  

   ' Add trailing backslash if missing

   sRoot = AddBackslash(sSpec)

   iCnt = 2& '.' '..'

  

   On Error Resume Next ' Ignore file errors

   sFile = Dir$(sRoot & "*.*", vbNormal)

   Do While LenB(sFile)

      SetAttributes sRoot & sFile, vbNormal

      Kill sRoot & sFile

      sFile = Dir$

   Loop

  

   On Error GoTo HandleIt ' No error should occur in here

   Do: sDir = Dir$(sRoot & "*", vbDirectory)

      For iIdx = 1& To iCnt

         sDir = Dir$ '.' '..' ['fail']

      Next

      If LenB(sDir) = 0& Then Exit Do

      If KillFolder(sRoot & sDir & DIR_SEP) Then

      ' Sub-folder is now gone but Dir$ was reset

      ' during recursive call so Do Dir$(..) again

      Else: iCnt = iCnt + 1&

      ' Kill folder failed (remnant files) so skip

      ' this folder (iCnt + 1) to get the rest

      End If

   Loop

  

   If bJustEmptyDontRemove = False Then RmDir sRoot ' Errors here if remnants

HandleIt:

   KillFolder = Not FolderExists(sSpec)

End Function

