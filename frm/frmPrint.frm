VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPrint 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1920
      ScaleHeight     =   525
      ScaleWidth      =   1245
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2310
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3690
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1245
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   525
      Left            =   180
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   926
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmPrint.frx":058A
   End
   Begin VB.Label lblPrint 
      Caption         =   "No Printer Discovered"
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   4785
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Print Form"
Option Explicit
'Parts of this thanks to ...
'NIXON Hyperwrite
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66376&lngWId=1

'and ... Microsoft KB article Q173981 and Q146022.


Private m_srtfText As String
Private Const c_Margin As Long = 450
Private moPrinter As Printer
Private mHavePrinter As Boolean
Private mPageCtn As Long


Private Type GETTEXTLENGTHEX
    flags As Long
    CodePage As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CHARRANGE
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As CHARRANGE ' Range of text to draw (see above declaration)
End Type

Private Type tPrintDlg
    
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
    
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type


Private Const WM_USER = &H400
Private Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Private Const GTL_PRECISE = 2
Private Const GTL_NUMBYTES = 16
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
'Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private gLeft As Long
Private gRight As Long
Private gTop As Long
Private gBottom As Long
'Private gHeader As String
'Private gFooter As String
Private mFormatRange As FormatRange
Private rectDrawTo As RECT
Private rectPage As RECT
Private TextLength As Long
Private newStartPos As Long
'Private dumpaway As Long

Private Type DEVMODE_TYPE
  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

' -------------------------------------------------------------------

' Windows API Declarations
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As tPrintDlg) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Function GetDefaultPrinter() As String
' Thanks to the coder who left this on PSC.
Dim def1 As String, def2 As String, def3 As String
Dim di As Long
    
   def2 = String(128, 0)
   ' Find default printer string
   di = GetPrivateProfileString("WINDOWS", "DEVICE", def1, def2, 127, def3)
   ' di = lenght of return string
   If di > 0 Then
      ' Parse string to printer name only using then comma Chr$(44)=","
      di = InStr(def2, Chr$(44)) - 1
      ' Test that di > 0
      If di Then GetDefaultPrinter = Left$(def2, di)
   End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Hide
'    modPrint.PrintRTF rtb, c_Margin, c_Margin, c_Margin, c_Margin
    PrintRTF rtb, c_Margin, c_Margin, c_Margin, c_Margin
    Unload Me
    Screen.MousePointer = vbDefault
End Sub

Public Property Let RTFText(ByVal srtfText As String)
    m_srtfText = srtfText
'    rtb.TextRTF = m_srtfText
End Property

Private Sub Form_Load()
Dim sP As String
Dim i As Long, j As Long
Dim lngTLen As Long
Dim lngPCount As Long
    If Len(Trim$(m_srtfText)) = 0 Then
        lblPrint.Caption = "Haven't anything to Print yet."
        Exit Sub
    End If
'    modPrint.WYSIWYG_RTF rtb, c_Margin, c_Margin, c_Margin, c_Margin, i, j
    WYSIWYG_RTF rtb, c_Margin, c_Margin, c_Margin, c_Margin, i, j
'    picPrint.Height = rtb.Height + (2 * c_Margin): picPrint.Width = rtb.Width + (2 * c_Margin)
    picPrint.Height = j: picPrint.Width = i
    
    If VB.Printers.Count > 0 Then
        sP = GetDefaultPrinter
        If Len(sP) Then
            For Each moPrinter In VB.Printers
                If moPrinter.DeviceName = sP Then
                    mHavePrinter = True
                    Exit For
                End If
            Next moPrinter
        End If
    End If
    Debug.Assert mHavePrinter           ' ... means no printer set.
    cmdPrint.Enabled = mHavePrinter
    If mHavePrinter = True Then
        rtb.TextRTF = m_srtfText
        lngTLen = GetLength(rtb.hwnd)
        lngPCount = PageCtnProc(lngTLen, rtb.hwnd, picPrint)
        If lngPCount = 0 And lngTLen > 0 Then lngPCount = 1
        lblPrint.Caption = "Print " & Format$(lngPCount, cNumFormat) & " page" & IIf(lngPCount <> 1, "s", "") & " to " & moPrinter.DeviceName & "?"
        lblPrint.Caption = lblPrint.Caption & vbNewLine & vbNewLine & "(page count is estimate)"
    End If
    ClearMemory
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' WYSIWYG_RTF - Sets an RTF control to display itself the same as it
'               would print on the default printer
'
' RTF - A RichTextBox control to set for WYSIWYG display.
'
' LeftMarginWidth - Width of desired left margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' Returns - The length of a line on the printer in twips
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)
   
Dim LeftOffSet As Long
Dim LeftMargin As Long
Dim RightMargin As Long
Dim TopOffSet As Long
Dim TopMargin As Long
Dim BottomMargin As Long
Dim printerhDC As Long
Dim r As Long
Dim lngHWnd As Long
    
    lngHWnd = RTF.hwnd
    
    Printer.ScaleMode = vbTwips
    ' Get the left offset to the printable area on the page in twips
    LeftOffSet = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
    LeftOffSet = Printer.ScaleX(LeftOffSet, vbPixels, vbTwips)
    
    ' Calculate the Left, and Right margins
    LeftMargin = LeftMarginWidth - LeftOffSet
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffSet
    
    ' Calculate the line width
    PrintableWidth = RightMargin - LeftMargin
    
    ' Get the top offset to the printable area on the page in twips
    TopOffSet = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
    TopOffSet = Printer.ScaleX(TopOffSet, vbPixels, vbTwips)
    
    ' Calculate the Left, and Right margins
    TopMargin = TopMarginWidth - TopOffSet
    BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffSet
    
    ' Calculate the line width
    PrintableHeight = BottomMargin - TopMargin
    
    
    ' Create an hDC on the printer pointed to by the printer object
    ' This DC needs to remain for the RTF to keep up the WYSIWYG display
    printerhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)
    
    ' Tell the RTF to base its display off of the printer
    ' at the desired line width
    ' r = SendMessage(RTF.hwnd, EM_SETTARGETDEVICE, printerhDC, ByVal PrintableWidth)
    r = SendMessage(lngHWnd, EM_SETTARGETDEVICE, printerhDC, ByVal PrintableWidth)

End Sub


Private Function GetLength(pRTBHwnd As Long, Optional bBytes As Boolean = False) As Long
' with thanks to Nixon Software.
Dim tGetLen As GETTEXTLENGTHEX
    tGetLen.CodePage = 0
    If bBytes = False Then
        tGetLen.flags = GTL_PRECISE
    Else
        tGetLen.flags = GTL_NUMBYTES
    End If
    GetLength = SendMessageAPI(pRTBHwnd, EM_GETTEXTLENGTHEX, tGetLen, 0)
End Function


' Test how many pages are there in total
Private Function PageCtnProc(pLen As Long, rtbHwnd As Long, inControl As Control) As Integer
' with thanks to Nixon Software.
Dim LeftOffSet As Long
Dim TopOffSet As Long
Dim LeftMargin As Long
Dim TopMargin As Long
Dim RightMargin As Long
Dim BottomMargin As Long
Dim r As Long
    
    ' Get the offset to the printable area on the page in twips.
    
    LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, 112), vbPixels, vbTwips)
    TopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, 113), vbPixels, vbTwips)
    
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = gLeft - LeftOffSet
    TopMargin = gTop - TopOffSet
    RightMargin = (Printer.Width - gRight) - LeftOffSet
    BottomMargin = (Printer.Height - gBottom) - TopOffSet

    ' Set printable area rect. Note in frmPrintPreview, scaleModes are all in pixels,
    ' have to compute the twips equivalent
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = Printer.ScaleWidth
    rectPage.Bottom = Printer.ScaleHeight
        
      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = LeftMargin
    rectDrawTo.Top = TopMargin
    rectDrawTo.Right = RightMargin
    rectDrawTo.Bottom = BottomMargin
  
      ' Set up the print instructions
    mFormatRange.hdc = inControl.hdc            ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = inControl.hdc      ' Point at hDC
    mFormatRange.rc = rectDrawTo                ' Area on page to draw to
    mFormatRange.rcPage = rectPage              ' Entire size of page
    mFormatRange.chrg.cpMin = 0                 ' Start of text
    mFormatRange.chrg.cpMax = -1                ' End of the text

    TextLength = pLen 'Len(frmMainText.Text1.Text)

    mPageCtn = 0
    
    Do
        ' ... Print the page by sending EM_FORMATRANGE message.
        newStartPos = SendMessage(rtbHwnd, EM_FORMATRANGE, True, mFormatRange)
        
        If newStartPos >= TextLength Then
'            If mPageCtn = 0 Then mPageCtn = 1
            Exit Do
        Else
            mPageCtn = mPageCtn + 1
        End If
        
        mFormatRange.chrg.cpMin = newStartPos       ' Starting position for next page
        mFormatRange.hdc = inControl.hdc
        mFormatRange.hdcTarget = inControl.hdc
    
    Loop
    
    inControl.Picture = LoadPicture()
    ' ... release the em_formatrange on the rtb.
    r = SendMessage(inControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    PageCtnProc = mPageCtn

End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins
'
' RTF - A RichTextBox control to print
'
' LeftMarginWidth - Width of desired left margin in twips
'
' TopMarginHeight - Height of desired top margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' BottomMarginHeight - Height of desired bottom margin in twips
'
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
'         parameter you should specify the same LeftMarginWidth and
'         RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
' with thanks to Nixon Software.
On Error GoTo 10
Dim LeftOffSet As Long, TopOffSet As Long
Dim LeftMargin As Long, TopMargin As Long
Dim RightMargin As Long, BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawto As RECT
Dim rcPage As RECT
Dim TextLength As Long
Dim NextCharPosition As Long
Dim r As Long
'Dim vPrintDlg As tPrintDlg
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE

Dim lpDevMode As Long, lpDevName As Long
Dim bReturn As Integer
Dim objPrinter As Printer, NewPrinterName As String
'Dim strSetting As String
Dim lngHWnd As Long
    
    lngHWnd = RTF.hwnd
    
    ' Get the offsett to the printable area on the page in twips
    LeftOffSet = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffSet
    TopMargin = TopMarginHeight - TopOffSet
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffSet
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffSet
    
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    
    ' Set rect in which to print (relative to printable area)
    rcDrawto.Left = LeftMargin
    rcDrawto.Top = TopMargin
    rcDrawto.Right = RightMargin
    rcDrawto.Bottom = BottomMargin
    
    Dim printDlg As tPrintDlg
     ' Set the starting information for the dialog box based on the current
     ' printer settings.
     
     printDlg.lStructSize = Len(printDlg)
     DevMode.dmDeviceName = Printer.DeviceName
     DevMode.dmSize = Len(DevMode)
     DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
     DevMode.dmOrientation = Printer.Orientation
     On Error Resume Next
     DevMode.dmDuplex = Printer.Duplex
     On Error GoTo 0
     
     ' Set the default PaperBin so that a valid value is returned even
     ' in the Cancel case.
    
     ' Set the flags for the PrinterDlg object using the same flags as in the
     ' common dialog control. The structure starts with VBPrinterConstants.
     'Allocate memory for the initialization hDevMode structure
     'and copy the settings gathered above into this memory
     printDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
        GMEM_ZEROINIT, Len(DevMode))
     lpDevMode = GlobalLock(printDlg.hDevMode)
     If lpDevMode > 0 Then
         CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
         bReturn = GlobalUnlock(lpDevMode)
     End If
     
     'Set the current driver, device, and port name strings
     With DevName
         .wDriverOffset = 8
         .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
         .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
         .wDefault = 0
     End With
     With Printer
         DevName.extra = .DriverName & Chr(0) & _
         .DeviceName & Chr(0) & .Port & Chr(0)
     End With
     
     'Allocate memory for the initial hDevName structure
     'and copy the settings gathered above into this memory
     printDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
         GMEM_ZEROINIT, Len(DevName))
     lpDevName = GlobalLock(printDlg.hDevNames)
     If lpDevName > 0 Then
         CopyMemory ByVal lpDevName, DevName, Len(DevName)
         bReturn = GlobalUnlock(lpDevName)
     End If
     
     'Call the print dialog up and let the user make changes
     If PrintDialog(printDlg) Then
     
         'First get the DevName structure.
         lpDevName = GlobalLock(printDlg.hDevNames)
             CopyMemory DevName, ByVal lpDevName, 45
         bReturn = GlobalUnlock(lpDevName)
         GlobalFree printDlg.hDevNames
     
         'Next get the DevMode structure and set the printer
         'properties appropriately
         lpDevMode = GlobalLock(printDlg.hDevMode)
             CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
         bReturn = GlobalUnlock(printDlg.hDevMode)
         GlobalFree printDlg.hDevMode
         NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
             InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
         If Printer.DeviceName <> NewPrinterName Then
             For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                     Set Printer = objPrinter
                End If
             Next
         End If
         On Error Resume Next
     
         'Set printer object properties according to selections made
         'by user
         With Printer
             .Copies = DevMode.dmCopies
             .Duplex = DevMode.dmDuplex
             .Orientation = DevMode.dmOrientation
             .ColorMode = DevMode.dmColor
             .PrintQuality = DevMode.dmPrintQuality
             .PaperSize = DevMode.dmPaperSize
         End With
         On Error GoTo 0
    Else
        'User chose Cancel
        Exit Sub
    End If
    
        ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips
     
    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawto            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    
    ' Get length of text in RTF
    TextLength = GetLength(lngHWnd) 'Must call function to workaround riched20 bug
    
    ' Loop printing each page until done
    Do
       ' Print the page by sending EM_FORMATRANGE message
       NextCharPosition = SendMessage(lngHWnd, EM_FORMATRANGE, True, fr)
       If NextCharPosition >= TextLength Then Exit Do 'If done then exit
       fr.chrg.cpMin = NextCharPosition ' Starting position for next page
       Printer.NewPage                  ' Move on to next page
       Printer.Print Space$(1) ' Re-initialize hDC
       fr.hdc = Printer.hdc
       fr.hdcTarget = Printer.hdc
    Loop
    
    ' Commit the print job
    Printer.EndDoc
    
    ' Allow the RTF to free up memory
    r = SendMessage(lngHWnd, EM_FORMATRANGE, False, ByVal CLng(0))
10:
End Sub





Private Sub Form_Unload(Cancel As Integer)
    ClearMemory
End Sub
