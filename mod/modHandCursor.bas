Attribute VB_Name = "modHandCursor"
Attribute VB_Description = "A module to help utislise the Windows Hand Cursor: includes Icon/Bitmap to Picture conversion function by handle (HandleToPicture)."
' what?
'  primarily written to get the windows hand cursor to
'  display as the mouse icon when moving over a label.
' why?
'  just because.
' when?
'  when using a label that can be clicked, e.g. a link.
' how?
'    ' ... module fields.
'    Private myHandCursor As StdPicture
'    Private myHand_handle As Long
'
'    ' ... get the hand cursor icon.
'    ' ... in a sub somewhere.
'    myHand_handle = modHandCursor.LoadHandCursor
'    If myHand_handle <> 0 Then
'        Set myHandCursor = modHandCursor.HandleToPicture(myHand_handle, False)
'        lblLink.MouseIcon = myHandCursor
'        lblLink.MousePointer = vbCustom
'    End If
'
'    The HandleToPicture method can be used to convert a handle to a picture
'    (bitmap or icon) and convert it to a picture.
'    The last parameter of this method describes whether the picture handle
'    relates to a bitmap (true) or icon (false).
' who?
'  LaVolpe: FYI, Hand Icon.
'  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63065&lngWId=1
'  I have taken what I needed.

Option Explicit

' -------------------------------------------------------------------
' ... Hand Cursor.
' ... Constant used to get hand cursor.
Private Const IDC_HAND As Long = 32649
' http://msdn.microsoft.com/en-us/library/ms648391%28v=VS.85%29.aspx
' ... the above link provides info on system cursors and their ids.

Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

' ... Used to convert icons/bitmaps to stdPicture objects.
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PictDesc, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long

' ... PICTDESC Type Declaration used with above declaration.
Private Type PictDesc
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type


Public Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture
Attribute HandleToPicture.VB_Description = "Convert an Icon/Bitmap Handle to a Picture object."

' ... convert an icon/bitmap handle to a Picture object

On Error GoTo ExitRoutine

Dim pic As PictDesc
Dim Guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    Guid(0) = &H7BF80980
    Guid(1) = &H101ABF32
    Guid(2) = &HAA00BB8B
    Guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, Guid(0), True, HandleToPicture

ExitRoutine:
End Function


Public Function LoadHandCursor() As Long
Attribute LoadHandCursor.VB_Description = "Try to get the Handle to the Hand Cursor for converting to a Picture."
' ... try to get the Hand Cursor as a picture.

    LoadHandCursor = LoadCursor(0, IDC_HAND)
    
End Function



