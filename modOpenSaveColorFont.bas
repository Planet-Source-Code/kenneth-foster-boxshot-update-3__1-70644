Attribute VB_Name = "modOpenSaveColorFont"
Option Explicit

' Original Author of this code is Mr.BoBo
'*******************************************************************

Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Private Const CF_NOSTYLESEL = &H100000
Const REGULAR_FONTTYPE = &H400
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Public Type CMDialog
    ownerform As Long
    Filter As String
    Filetitle As String
    Filterindex As Long
    Filename As String
    Initdir As String
    Dialogtitle As String
    Flags As Long
End Type

Public Type SFfont
    FontName As String
    Fontsize As Integer
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    FontStrikethru As Boolean
    FontColor As Long
End Type

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long

Public SelectFont As SFfont
Public cmnDlg As CMDialog

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim OFName As OPENFILENAME
Dim LastDir As String

Public Function ShowOpen(frm As Form) As String
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hwndOwner = frm.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    OFName.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = LastDir
    'Set the dialog title
    OFName.lpstrTitle = "Open File "
    'no extra flags
    OFName.Flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function

Public Function ShowFont(Optional NoFX As Boolean) As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim retval As Long
    lfont.lfHeight = 0
    lfont.lfItalic = SelectFont.FontItalic
    lfont.lfUnderline = SelectFont.FontUnderline
    lfont.lfStrikeOut = SelectFont.FontStrikethru
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfHeight = SelectFont.Fontsize * 1.33
    If SelectFont.FontBold Then
        lfont.lfWidth = 700
    Else
        lfont.lfWidth = 0
    End If
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = SelectFont.FontName & vbNullChar
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = cmnDlg.ownerform
    cf.lpLogFont = pMem
    cf.iPointSize = SelectFont.Fontsize * 10
    cf.rgbColors = SelectFont.FontColor
    If SelectFont.FontBold Or SelectFont.FontItalic Then
        cf.nFontType = IIf(SelectFont.FontBold, BOLD_FONTTYPE, 0) Or IIf(SelectFont.FontItalic, ITALIC_FONTTYPE, 0)
    Else
        cf.nFontType = REGULAR_FONTTYPE
    End If
    cf.nSizeMin = 10
    cf.nSizeMax = 72
    If NoFX Then
        cf.Flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT  'Or CF_USESTYLE
    Else
        cf.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT 'Or CF_USESTYLE
    End If
    retval = CHOOSEFONT(cf)
    If retval <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        With SelectFont
            .FontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
            .FontBold = False
            .FontItalic = False
            .FontUnderline = False
            .FontStrikethru = False
            .Fontsize = cf.iPointSize / 10
            If lfont.lfWeight = 700 Then .FontBold = True
            .FontItalic = lfont.lfItalic
            .FontUnderline = lfont.lfUnderline
            .FontStrikethru = lfont.lfStrikeOut
            .FontColor = cf.rgbColors
        End With
    Else
        ShowFont = False
    End If
    retval = GlobalUnlock(hMem)
    retval = GlobalFree(hMem)
End Function

