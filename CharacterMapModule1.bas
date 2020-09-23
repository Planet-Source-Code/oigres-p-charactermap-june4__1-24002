Attribute VB_Name = "Module1"
Option Explicit
'From MSDN article
'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Public mousevisible As Boolean
Type RECT   '  16  Bytes
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, _
        lprect As RECT)

Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)

Type LOGFONT
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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias _
        "EnumFontFamiliesA" _
        (ByVal hdc As Long, ByVal lpszFamily As String, _
        ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
        ByVal hdc As Long) As Long
Private m_selectedsquare


Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
            ByVal FontType As Long, LParam As ListBox) As Long
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    LParam.AddItem left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function

Sub FillListWithFonts(LB As ComboBox) 'ListBox)
    Dim hdc As Long
    LB.Clear
    hdc = GetDC(LB.hWnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
    ReleaseDC LB.hWnd, hdc
End Sub

'/******************************************************************************
Function Bin(ByVal value As Long, Optional digits As Long = -1) As String
'/******************************************************************************

    Dim result As String, exponent As Integer
    ' this is faster than creating the string by appending chars
    result = String$(32, "0")
    Do
        If value And Power2(exponent) Then
            ' we found a bit that is set, clear it
            Mid$(result, 32 - exponent, 1) = "1"
            value = value Xor Power2(exponent)
        End If
        exponent = exponent + 1
    Loop While value
    If digits < 0 Then
        ' trim non significant digits, if digits was omitted or negative
        Bin = Mid$(result, 33 - exponent)
    Else
        ' else trim to the requested number of digits
        Bin = right$(result, digits)
    End If
End Function

' Raise 2 to a power
' the exponent must be in the range [0,31]

'/******************************************************************************
Function Power2(ByVal exponent As Long) As Long
'/******************************************************************************
    Static res(0 To 31) As Long
    Dim i As Long

    ' rule out errors
    If exponent < 0 Or exponent > 31 Then Err.Raise 5

    ' initialize the array at the first call
    If res(0) = 0 Then
        res(0) = 1
        For i = 1 To 30
            res(i) = res(i - 1) * 2
        Next
        ' this is a special case
        res(31) = &H80000000
    End If

    ' return the result
    Power2 = res(exponent)

End Function

Property Let selectedsquare(v As Long)
    m_selectedsquare = v
End Property
Property Get selectedsquare() As Long
    selectedsquare = m_selectedsquare
End Property
'/******************************************************************************
Public Sub cboClr(cb As ComboBox)
'/******************************************************************************

    Dim x, y, clr As Long, r, g, b, index As Long
    For x = 0 To 35
        For y = 0 To 5
            clr = RGB(r, g, b)
            'Form1.Line (x * BlockWidth, y * BlockWidth)-(x * BlockWidth + BlockWidth - 1, y * BlockWidth + BlockWidth - 1), clr, BF
cb.AddItem clr
            index = index + 1
            r = r + 51
            If r > 255 Then
                r = 0: g = g + 51
            End If
            If g > 255 Then
                g = 0: b = b + 51
            End If
            If b > 255 Then
                b = 0:
            End If
        Next

    Next
End Sub

'Public Sub myShowCursor(ByVal bdisplay As Boolean)
'Dim ret As Long
''hide cursor
'If bdisplay = False Then
'ret = ShowCursor(False)
'        'showcursor shows cursor if the return count >=0
'        'force it to hide
'        While ret >= 0
'            ret = ShowCursor(False)
'        Wend
'End If
'If bdisplay = True Then
'
'
'End If

'End Sub
Public Sub makeCursorInvisible()
    '/******************************************************************************
    Dim ret As Long
    ret = ShowCursor(False)
    While ret >= 0
        ret = ShowCursor(False)
    Wend
    mousevisible = False
End Sub
Public Sub makeCursorVisible()
    '/******************************************************************************
    Dim ret As Long
    ret = ShowCursor(True)
    While ret < 0
        ret = ShowCursor(True)
    Wend
    mousevisible = True
End Sub

Public Function fileExists(str As String) As Boolean
    fileExists = Dir$(str) <> ""
End Function
