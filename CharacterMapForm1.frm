VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Map"
   ClientHeight    =   4200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "CharacterMapForm1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSample 
      Height          =   315
      Left            =   6600
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtrealClr 
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtNameHex 
      Height          =   315
      Left            =   5520
      TabIndex        =   34
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cboHtmlClrName 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   6
      Left            =   3240
      TabIndex        =   32
      Top             =   2640
      Width           =   135
   End
   Begin VB.TextBox txtHtmlName 
      Height          =   315
      Left            =   3360
      TabIndex        =   31
      ToolTipText     =   "HtmlNamedEntity"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   5
      Left            =   4680
      TabIndex        =   30
      Top             =   3015
      Width           =   135
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   4
      Left            =   4680
      TabIndex        =   29
      Top             =   2640
      Width           =   135
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   3
      Left            =   1920
      TabIndex        =   28
      Top             =   2640
      Width           =   135
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   3360
      Width           =   135
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   26
      Top             =   3000
      Width           =   135
   End
   Begin VB.CommandButton cmdCopyPush 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   2640
      Width           =   135
   End
   Begin VB.ComboBox cboHtmlClr 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtHtmlClr 
      Height          =   315
      Left            =   4800
      TabIndex        =   23
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txthtml 
      Height          =   315
      Left            =   2040
      TabIndex        =   22
      ToolTipText     =   "Html Entity"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6480
      Top             =   120
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6600
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Txt3 
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Height          =   315
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   5
      Top             =   720
      Width           =   6255
   End
   Begin VB.ComboBox CboFonts 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Txtcopy 
      Height          =   375
      HideSelection   =   0   'False
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   6480
      Max             =   100
      Min             =   1
      TabIndex        =   20
      Top             =   2040
      Value           =   50
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "&Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Char&acters to copy:"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   1920
      TabIndex        =   17
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Dec:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Bin:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   3840
      Width           =   2820
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuWhatsThis 
         Caption         =   "&what's this?"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'******************************************************************************
'*Character Map recreation -17/jul/2000
'******************************************************************************
''improved' 16/3/2001
'By oigres P Email:oigres@postmaster.co.uk
Private Type POINTAPI  '  8 Bytes
    x As Long
    y As Long
End Type
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
        ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As _
        Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As _
        Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
        ByVal y As Long, lpPoint As POINTAPI)
Private Declare Function CreateRectRgnIndirect& Lib "gdi32" (lprect As RECT)
Private Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long)
''Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal _
        y As Long)
Private Declare Function Rectangle& Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Const HORZRES = 8
Private Const VERTRES = 10
Const SRCCOPY = &HCC0020
Dim HtmlNames(255) As String
Dim HtmlClrHex(139) As String
Dim asciiList() ' list of character descriptions
Dim sizeX, sizeY, previousX, previousY
Dim mouseDown As Boolean, bDrawLine As Boolean
Dim mode As Long, loadokay As Boolean
Private Sub CboFonts_Click()

    drawSquare CboFonts.List(CboFonts.ListIndex)
    Picture2.Font = CboFonts.List(CboFonts.ListIndex)
    Picture2.FontSize = 18
    Txtcopy.Font = CboFonts.List(CboFonts.ListIndex)

    'reselect last square
    drawfocusColour previousX, previousY


End Sub

Private Sub cboHtmlClr_Change()
MsgBox "hello"
cboHtmlClr.ToolTipText = "hello" 'cboHtmlClr.ListIndex
End Sub

Private Sub cboHtmlClr_Click()
Dim clr As Long, hstr As String 'hexdecimal string
Dim rpercent, gpercent, bpercent
clr = Val(cboHtmlClr.List(cboHtmlClr.ListIndex))
Dim lR As Long, lG As Long, LB As Long

    lR = (clr Mod &H100)
    lG = (clr \ &H100) Mod &H100
    LB = (clr \ &H10000) Mod &H10000


txtHtmlClr.BackColor = clr

txtHtmlClr.ForeColor = (16 ^ 6) - (clr + 1)
hstr = hstr & Format(Hex(lR), "00")
hstr = hstr & Format(Hex(lG), "00")
hstr = hstr & Format(Hex(LB), "00")
cboHtmlClr.ToolTipText = "Colour number " & cboHtmlClr.ListIndex + 1 & ":Hex " & hstr

rpercent = Int((lR / 255) * 100)
gpercent = Int((lG / 255) * 100)
bpercent = Int((LB / 255) * 100)

txtHtmlClr.Text = "#" & hstr
txtHtmlClr.ToolTipText = rpercent & "%:" & gpercent & "%:" & bpercent & "%"  'Hex(clr) & ":" & Hex(lR) & Hex(lG) & Hex(lB) & ":" & hstr


End Sub

Private Sub cboHtmlClr_KeyPress(KeyAscii As Integer)
'cboHtmlClr.ToolTipText = cboHtmlClr.ListIndex
End Sub

Private Sub cboHtmlClrName_Click()
    Dim temp As Long, rghstr As String, rstr As String
    Dim gstr As String, bstr As String
    Dim lresult As Long, r1, g1, b1, tmpRealClr As Long
    Me.txtNameHex.Text = HtmlClrHex(cboHtmlClrName.ListIndex)
    rghstr = right$(Me.txtNameHex.Text, Len(Me.txtNameHex.Text) - 1)
    rstr = Mid$(rghstr, 1, 2)
    r1 = (Val("&h" & rstr) \ 51) * 51
    gstr = Mid$(rghstr, 3, 2)
    g1 = (Val("&h" & gstr) \ 51) * 51
    bstr = Mid$(rghstr, 5, 2)
    b1 = (Val("&h" & bstr) \ 51) * 51
    Debug.Print r1; Hex(r1); g1; Hex(g1); b1; Hex(b1)
    lresult = RGB(r1, g1, b1)
    temp = RGB(Val("&h" & rstr), Val("&h" & gstr), Val("&h" & bstr))
    'temp = CDbl(Val("&h" & rghstr & "&")) 'need long value in string to val
    Me.txtNameHex.BackColor = temp
    Me.txtNameHex.ForeColor = (16 ^ 6) - (temp + 1)
    'lresult = temp \ 51
    'Debug.Print "1: " & lresult & ":" & Hex(lresult)
    'lresult = lresult * 51
    'Debug.Print "2: " & lresult & ":" & Hex(lresult)

    'Me.txtNameHex.ToolTipText = "#" & Format(Hex(r1), "00") & Format(Hex(g1), "00") & Format(Hex(b1), "00")
    Me.txtNameHex.ToolTipText = "Named Colour Value"
    tmpRealClr = RGB(r1, g1, b1)
    Me.txtrealClr.BackColor = tmpRealClr
    'try to get matching value in cbohtmlclr
    Dim x As Integer
    cboHtmlClr.Visible = False
    For x = 0 To cboHtmlClr.ListCount - 1
        If cboHtmlClr.List(x) = CStr(tmpRealClr) Then
            cboHtmlClr.ListIndex = x
            Exit For
        End If
        
    Next x
    'cboHtmlClr.Text = tmpRealClr
    cboHtmlClr.Visible = True
    txtrealClr.ToolTipText = "Web Colour:" & "#" & Format(Hex(r1), "00") & Format(Hex(g1), "00") & Format(Hex(b1), "00")
    'Me.txtNameHex.ToolTipText = "#" & Hex(lresult)

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    'copy to clipboard
    Clipboard.Clear
    Clipboard.SetText Txtcopy.Text, vbCFText
    Picture1.SetFocus
End Sub



Private Sub cmdCopyPush_Click(index As Integer)
    With Clipboard
        .Clear

        Select Case index

        Case 0
            .SetText Txt1.Text
        Case 1
            .SetText Txt2.Text
        Case 2
            .SetText Txt3.Text
        Case 3
            .SetText txthtml.Text
        Case 4
            .SetText txtHtmlClr.Text
        Case 5
            .SetText cboHtmlClr.List(cboHtmlClr.ListIndex)
        Case 6
.SetText txtHtmlName.Text
        End Select
    End With
End Sub

Private Sub cmdSelect_Click()
    '
    inserttext
    Picture1.SetFocus
End Sub
Sub inserttext()
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety, s
    s = selectedsquare
    Y1 = s \ 32
    X1 = s Mod 32
    char$ = Chr$((Y1 * 32) + (X1 + 1) + 30) '1)
    Txtcopy.SelText = char$

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'************************************************************
'* Name: Form_KeyDown
'*
'* Description :
'* Parameters :
'* Created : 17-Jul-2000
'************************************************************

    
    ''MsgBox "keydon " & KeyCode & ":" & ActiveControl
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        MsgBox "alt+ A=frm key"
        'Txtcopy.SelText =
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If



    If KeyCode = Asc("F") And (Shift And vbAltMask) Then
        'MsgBox "alt+ A"
        'Txtcopy.SelText =
        CboFonts.SetFocus
    End If
    If KeyCode = Asc("S") And (Shift And vbAltMask) Then
        MsgBox "alt+ S"
        'Txtcopy.SelText =
        'CboFonts.SetFocus
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        ''MsgBox "alt+ A=frm key"
        'Txtcopy.SelText =
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If
    If KeyCode = Asc("F") And (Shift And vbAltMask) Then
        'MsgBox "alt+ A"
        'Txtcopy.SelText =
        CboFonts.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim x, y, value, lastleft, lasttop
    Dim index
    '' Form1.ScaleWidth = 32 * 7
    loadokay = True
    bDrawLine = True 'magnifier cross is on
    Form1.Icon = LoadPicture(App.Path & "\charmap.ico")
    sizeX = (Picture1.ScaleWidth \ 32) ' + 1 '  32*7=224
    sizeY = (Picture1.ScaleHeight \ 7) ''' + 1 '  32*7=224
    '''MsgBox sizeX & ":" & sizeY
    
    'load info files
    
    If createAsciiList() = True Then
        If createHtmlNamed() = True Then
            If createHtmlColorName() = True Then
            
            Else
                loadokay = False
            End If
        Else
            loadokay = False
        End If
    Else
        loadokay = False
    End If
    If loadokay = False Then
        MsgBox "Errors on loading"
        Unload Me
        'need to exit sub to stop code after executing
        Exit Sub
    End If
    
    '
    
    
    
    Me.cboHtmlClrName.ListIndex = 0

    CboFonts.Visible = False
    FillListWithFonts CboFonts 'List1
    
    'load and select combo htmlclr
    cboClr Me.cboHtmlClr
    cboHtmlClr.ListIndex = 0
    'getlast setting if saved
    value = GetSetting("MyCharacterMap", "CboFonts", "LastFont", "0")
    ''cbofonts.l
    CboFonts.ListIndex = value
    'get last font
    value = GetSetting("MyCharacterMap", "CboFonts", "LastFontName", "Times New Roman")
    'somehow we saved a zerolength string (registry has an entry of "")
    If value = "" Then value = "Times New Roman"
    drawSquare CStr(value)
    
    
    CboFonts.Visible = True

    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
    Txtcopy.Text = GetSetting("MyCharacterMap", "Txt", "LastText", "")
    Txtcopy.SelStart = Len(Txtcopy.Text)
    Txtcopy_Change
    
    'starts off with first square selected
    Dim xp, yp
    
    selectedsquare = GetSetting("MyCharacterMap", "Form1", "selectedsquare", "1")
    'bug->boundary error (was selectedsquare Mod 32), error on value 224
    'bug area --------------------------------------------
    'Dim myselsqr As Long
    'myselsqr = selectedsquare - 1
    '
    xp = ((selectedsquare - 1) Mod 32) ' (224-1) mod 32 = 31
    'xp = (myselsqr Mod 32)
    ''MsgBox xp
    yp = (selectedsquare - 1) \ 32 '(224-1) \ 32 =6
    ''MsgBox yp
    ''MsgBox "selectedsquare = " & selectedsquare
    drawselected selectedsquare - 1
   
    'bug propagated to here - was xp * sizex + 13
    Picture1_MouseDown 0&, 0&, xp * sizeX, yp * sizeY + 4  '(selectedsquare Mod 32) * 32, (selectedsquare \ 32) * 32
    Picture1_MouseUp 0&, 0&, xp * sizeX, yp * sizeY + 4  '(selectedsquare Mod 32) * 32, (selectedsquare \ 32) * 32
    
    'get last form coords
    lastleft = GetSetting("MyCharacterMap", "Form1", "Left", "0")
    If lastleft < 0 Then lastleft = 0
    lasttop = GetSetting("MyCharacterMap", "Form1", "Top", "0")
    If lasttop < 0 Then lasttop = 0
    
    value = GetSetting("MyCharacterMap", "Magnification", "Last", "50")
    If value < 1 Or value > 100 Then value = 50
    VScroll1.value = value
    
    Form1.Move lastleft, lasttop
    Form1.Show
    Picture1.SetFocus
   
End Sub

'''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub updateLabel(x, y)
    Dim key, k$
    'give keystroke and alt information
    'changed 11 june 2001 -was (x+1)
    'changed back to (x+1)
    key = (y * 32) + (x + 1) ' + 31
    k$ = "Keystroke: "
    'MsgBox key
    Select Case key
    Case 1
        Label4.Caption = k$ & "Spacebar"
    Case 2 To 95 '
        If key = 7 Then 'need && to show in label for ampersand
            Label4.Caption = k$ & "&&" 'Chr$(key + 31)
        Else
            Label4.Caption = k$ & Chr$(key + 31)
        End If

    Case 96 To 97
        Label4.Caption = k$ & "Ctrl+" & (key - 95)
    Case 98 To 224
        Label4.Caption = k$ & "Alt+0" & key + 31
    End Select
    'hex / bin text
    Txt1.Text = Hex(key + 31)
    Txt2.Text = Bin(key + 31, 8)
    Txt3.Text = key + 31
    'bug if I closed prog with last select character as 'ÿ'
    '225+31 =256;started at form load-
    'Debug.Assert key + 31 < 256
    If key + 31 < 256 Then
        txtHtmlName.Text = HtmlNames(key + 31)
    End If
    Label1.Caption = "Col: " & x & " Line: " & y & " Square:" & (y * 32) + (x + 1) & " Ascii: " & key + 31 ' * (y1 + 1)
    'Debug.Print key
    'asciilist array starts at 0 index
    Select Case key
    Case 1 To 98
        Label7.Caption = asciiList(key - 1)
    Case 99 To 129
        Label7.Caption = asciiList(key - 1)
    Case 130 To 224
        Label7.Caption = asciiList(key - 1)

    End Select
'update html entity
    txthtml.Text = "&#" & key + 31 & ";"
'update colour square
'txtHtmlClr.BackColor = 77672 * (216 Mod key)
'txtHtmlClr.ForeColor = (16 ^ 6) - (77672 * (216 Mod key))
'txtHtmlClr.Text = Hex(txtHtmlClr.BackColor)

End Sub
Function createAsciiList() As Boolean
    'assume true
    createAsciiList = True
    ReDim asciiList(250)
    Dim a$, index As Long, ffile As Long
    ffile = FreeFile()
    If fileExists(App.Path & "\asciiquoteds.txt") Then
    
        Open App.Path & "\asciiquoteds.txt" For Input As ffile
        Do While Not (EOF(ffile))
            Input #ffile, a$
            asciiList(index) = a$
            index = index + 1
        Loop
    
    
        Close ffile
    Else
        Close ffile
        createAsciiList = False
        MsgBox "File " & App.Path & "\asciiquoteds.txt" & " not found"
        
    End If
   
End Function

Function createHtmlNamed() As Boolean
 Dim a$, index As Long, ffile As Long
    createHtmlNamed = True 'assume success
    ffile = FreeFile()
    If fileExists(App.Path & "\htmlentand.txt") Then
        Open App.Path & "\htmlentand.txt" For Input As ffile
            Do While Not (EOF(ffile))
                Input #ffile, a$
                HtmlNames(index) = a$
                index = index + 1
            Loop
            '''MsgBox "create HtmlNamed index=  " & index - 1 ' -1 because it is incremented
        
        Close ffile
    Else
        Close ffile
        MsgBox "File " & App.Path & "\htmlentand.txt" & " not found"
        createHtmlNamed = False
        'Unload Me
    End If
End Function

Function createHtmlColorName() As Boolean
Dim a$, index As Long, ffile As Long
    createHtmlColorName = True
    ffile = FreeFile()
    If fileExists(App.Path & "\htmlcolournames.txt") Then
    Open App.Path & "\htmlcolournames.txt" For Input As ffile
    Do While Not (EOF(ffile))
        Input #ffile, a$
        Me.cboHtmlClrName.AddItem a$
        Input #ffile, a$
        HtmlClrHex(index) = a$
        index = index + 1
    Loop
    Close ffile
Else
    Close ffile
    MsgBox "File " & App.Path & "\htmlcolournames.txt"
    createHtmlColorName = False
End If

    
End Function


Private Sub Form_Unload(Cancel As Integer)

If loadokay = True Then
    SaveSetting "MyCharacterMap", "CboFonts", "LastFont", CboFonts.ListIndex
    SaveSetting "MyCharacterMap", "CboFonts", "LastFontName", CboFonts.List(CboFonts.ListIndex)
    SaveSetting "MyCharacterMap", "Form1", "Left", Form1.left
    SaveSetting "MyCharacterMap", "Form1", "Top", Form1.top
    SaveSetting "MyCharacterMap", "Magnification", "Last", Form1.VScroll1.value
    'poosible bug
    SaveSetting "MyCharacterMap", "Form1", "selectedSquare", Module1.selectedsquare
    SaveSetting "MyCharacterMap", "Txt", "LastText", Txtcopy.Text
End If
End Sub

Private Sub mnuWhatsThis_Click()
    MsgBox "By oigres P", , "What's this?"
End Sub

Private Sub Picture1_DblClick()
    inserttext
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDown
        If selectedsquare + 32 < 225 Then
            selectedsquare = selectedsquare + 32
        End If
    Case vbKeyUp
        If selectedsquare - 32 > 0 Then
            selectedsquare = selectedsquare - 32
        End If
    Case vbKeyRight
        If selectedsquare + 1 < 225 Then
            selectedsquare = selectedsquare + 1
        End If
    Case vbKeyLeft
        If selectedsquare - 1 > 0 Then
            selectedsquare = selectedsquare - 1
        End If
    Case Else
        Exit Sub
    End Select
    drawselected (selectedsquare - 1)
    updateLabel (selectedsquare - 1) Mod 32, (selectedsquare - 1) \ 32

End Sub
'/******************************************************************************
Sub drawselected(s As Long)
    '/******************************************************************************
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety
    ''MsgBox "draw selected input " & s
    Y1 = s \ 32
    X1 = s Mod 32
    
    
    'erase previous ?
    Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
    Picture1.CurrentX = (previousX * sizeX) + 3
    Picture1.CurrentY = (previousY * sizeY)

    Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
    previousX = X1
    previousY = Y1

    char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = False: Picture3.Visible = False
    offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
    offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
    Picture2.left = (X1 * sizeX - 5) + 10
    Picture2.top = (Y1 * sizeY - 5) + 35
    Picture3.left = Picture2.left + 5
    Picture3.top = Picture2.top + 5
    Picture2.CurrentX = offsetx
    Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
    Picture2.Picture = LoadPicture()
    Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = True: Picture3.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseDown
    '*
    '* Description:
    '*
    '* Date Created:  7/17/00
    '*
    '* Created By: oigres P
    '*
    '* Modified: 7/19/00
    '*
    '*******************************************************************************
    Dim X1, Y1, ret, lprect As RECT, offsetx, offsety, char$
    X1 = x \ sizeX
    Y1 = y \ sizeY
    Debug.Print X1 & ":" & Y1
    If Button = vbRightButton Then
        Form1.PopupMenu mnuFile
        Exit Sub
    End If
    'if in square of picture
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
        ''If x1 <> previousX And y <> previousY Then
        'erase previous focus rectangle
        ''MsgBox IsEmpty(previousX)
        If Not (IsEmpty(previousX) And IsEmpty(previousY)) Then
            lprect.left = X1 * sizeX + 1
            lprect.top = Y1 * sizeY + 1
            lprect.right = X1 * sizeX + (sizeX - 1) + 1 '- 1
            lprect.bottom = Y1 * sizeY + (sizeY - 1) + 1
            ''DrawFocusRect Picture1.hdc, lprect

            Picture1.Line (previousX * sizeX, previousY * sizeY)-(previousX * sizeX + (sizeX), previousY * sizeY + (sizeY)), vbBlack, BF
            Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
           
            char$ = Chr$((previousY * 32) + (previousX + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (previousX * sizeX) + offsetx
            Picture1.CurrentY = (previousY * sizeY) + offsety
            Picture1.Print char$;
        End If
        Picture2.Visible = False
        Picture3.Visible = False
        Picture2.left = (X1 * sizeX - 5) + 10
        Picture2.top = (Y1 * sizeY - 5) + 35
        Picture3.left = Picture2.left + 5
        Picture3.top = Picture2.top + 5
        Picture2.Visible = True
        Picture3.Visible = True
        selectedsquare = (Y1 * 32) + (X1 + 1)

        previousX = X1
        previousY = Y1
    
    End If ' in square


    Call updateLabel(X1, Y1)
    
    'hide cursor
    If mouseDown = False Then
        
        makeCursorInvisible
      
        mousevisible = False
        '' Label5.Caption = "showcursor times= " & ret
    End If
    mousevisible = False
    ''Form1.MousePointer = 15
    Picture2.Visible = True
    Picture3.Visible = True
    mouseDown = True
End Sub

'/******************************************************************************
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseMove
    '*
    '* Description:
    '*
    '* Date Created:  7/21/00
    '*
    '* Created By:
    '*
    '* Modified:
    '*
    '*******************************************************************************

    Dim X1, Y1, ret, char$, key
    Dim offsetx, offsety
    Static lastx
    Static lasty

    X1 = x \ sizeX
    Y1 = y \ sizeY

    If mouseDown = True Then
        If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
            If mousevisible = True Then
                makeCursorInvisible
            End If
            If lastx = X1 And lasty = Y1 Then Exit Sub
            lastx = X1: lasty = Y1
            key = (Y1 * 32) + (X1 + 1)

            Picture2.Visible = False
            Picture3.Visible = False
            Picture2.left = (X1 * sizeX - 5) + 10
            Picture2.top = (Y1 * sizeY - 5) + 35
            Picture3.left = Picture2.left + 5
            Picture3.top = Picture2.top + 5
            '            Picture2.Visible = True
            '            Picture3.Visible = True
            char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
            If Picture2.Tag = char$ Then
            Else

                '        Picture1.Picture = LoadPicture()
                '        Picture1.CurrentX = 0: Picture1.CurrentY = 0
                '        Picture1.Print Chr$((y1 * 32) + (x1 + 1) + 31)
                previousX = X1
                previousY = Y1
                ''Picture2.Visible = False
                Picture2.Tag = char$

                offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
                offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
                Picture2.CurrentX = offsetx
                Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
                Picture2.Picture = LoadPicture()
                Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
                Picture2.Visible = True
                Picture3.Visible = True
            End If 'if tag

            Call updateLabel(X1, Y1)
            previousX = X1
            previousY = Y1
        Else 'not in square
            'showcursor ?
            makeCursorVisible


            Exit Sub
        End If ' x1 >= 0 And x1 <= 31 And y1 >= 0 And y1 <= 6
        
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseUp
    '*
    '* Description:
    '*
    '* Date Created:  7/21/00
    '*
    '* Created By:
    '*
    '* Modified:
    '*
    '*******************************************************************************
    Dim ret, X1, Y1, lprect As RECT
    X1 = x \ sizeX
    Y1 = y \ sizeY
    If mousevisible = False Then
        makeCursorVisible

        mousevisible = True
    End If
    
    drawfocusColour previousX, previousY
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then

    Else
        If mouseDown = True Then
            Picture2.Visible = False
            Picture3.Visible = False
            'draw focus rectangle
            drawfocusColour previousX, previousY
        End If

    End If
    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
End Sub

'/******************************************************************************
Sub drawSquare(f As String)
    '/******************************************************************************
    'draw the font characters in the grid (picture1)
    Dim x As Long, y As Long, char$, lpPT As POINTAPI
    Dim offsetx, offsety
    Picture1.Visible = False
    Picture1.FontName = f
    Picture1.FontSize = 8
    Picture1.Picture = LoadPicture() 'fast clear
    For x = 0 To 31 '32
        For y = 0 To 6 '7
            ''Picture1.Line (x * sizex, y * sizey)-(x * sizex + (sizey - 1), y * sizex + (sizey - 1)), vbBlack, B
            char$ = Chr$((y * 32) + (x + 1) + 31)
            'centre the character in the grid square
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (x * sizeX) + offsetx
            Picture1.CurrentY = (y * sizeY) + offsety
            Picture1.Print char$;

        Next y
    Next x
    'draw grid lines
    For x = 0 To 7
        MoveToEx Picture1.hdc, 0, x * sizeY, lpPT
        LineTo Picture1.hdc, sizeX * 32, x * sizeY
    Next x
    For x = 0 To 32
        MoveToEx Picture1.hdc, x * sizeX, 0, lpPT
        LineTo Picture1.hdc, x * sizeX, sizeY * 7 + 1 'Picture1.ScaleHeight - 1
    Next x
    Picture1.Visible = True
End Sub


Private Sub Picture4_Click()

'bDrawLine = Not bDrawLine

End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbKeyLButton Then
bDrawLine = Not bDrawLine

End If
'If Button = vbKeyRButton Then
''bDrawLine = Not bDrawLine
'mode = mode + 1
'If mode > 16 Then
'    mode = 1
'End If
'Picture4.DrawMode = mode
'End If

End Sub

Private Sub Timer1_Timer()
    Dim cp As POINTAPI, hr As Long, vr As Long, ret As Long
Dim clr As Long
    Static lastcpx
    Static lastcpy
    GetCursorPos cp
    ''Label1.Caption = cp.x & Space(6 - Len(CStr(cp.x))) & ":" & cp.y

    Dim dsDC As Long, lpPT As POINTAPI, dshwnd As Long, Percent
    Dim lengthx, lengthy, offsetx, offsety, blitareax, blitareay
    'get desktop device context
    dsDC = GetDC(0&)
    'get screen width, height
    hr = GetDeviceCaps(dsDC, HORZRES)
    vr = GetDeviceCaps(dsDC, VERTRES)

    dshwnd = GetDesktopWindow()
    '      vscroll1=1..100 so 1/100=.1; 100/100=1;New Resolution
    Percent = VScroll1.value / 100
    lengthx = (Picture4.ScaleWidth - 0) * Percent
    lengthy = (Picture4.ScaleHeight - 0) * Percent
    'center image about mouse
    offsetx = lengthx \ 2
    offsety = lengthy \ 2
    blitareax = Picture4.ScaleWidth - 0 'actual area to blit to
    blitareay = Picture4.ScaleHeight - 0
    
    
    
    'Debug.Print lengthx; lengthy; Percent; offsetx; offsety
    'stop copying the screen off the edges <0 and  >horzres
    If cp.x - offsetx >= 0 And cp.x + offsetx < hr Then '800=screen width
        lastcpx = cp.x
    End If
    If cp.y - offsety >= 0 And cp.y + offsety < vr Then '600= screen height
        lastcpy = cp.y
            '                dest hdc ,destx,desty,width,height, sourceDC, source x,sourcey,sourcewidth,sourceheight,raster operation
    End If
    ret = StretchBlt(Picture4.hdc, 0, 0, blitareax, blitareay, dsDC, lastcpx - offsetx, lastcpy - offsety, lengthx, lengthy, SRCCOPY)
    clr = GetPixel(dsDC, cp.x, cp.y)
    If clr > -1 Then
    If bDrawLine = True Then
    Picture4.Line (0, 0)-(Picture4.Width - 1, Picture4.Height - 1) ', (16 ^ 6) - (clr + 1)
    Picture4.Line (Picture4.Width - 3, 0)-(0, Picture4.Height - 1) ', (16 ^ 6) - (clr + 1)
    End If
    ''update colour under cursor to picbox - picSample
    
    
    picSample.BackColor = clr
    picSample.ToolTipText = clr
    picSample.CurrentX = 0: picSample.CurrentY = 0
    picSample.ForeColor = (16 ^ 6) - (clr + 1)
    
    picSample.Print clr
    
    End If
    
    'Form1.Line (0, 0)-(Form1.ScaleWidth - VScroll1.Width, Form1.ScaleHeight - Label1.Height)
    'Form1.Line (Form1.ScaleWidth - VScroll1.Width, 0)-(0, Form1.ScaleHeight - Label1.Height)
    ReleaseDC dshwnd, dsDC 'previous bug not releasing memory
    Label5.Caption = Format(100 / VScroll1.value, "FIXED") & ":" & cp.x & ":" & cp.y




End Sub

'/******************************************************************************
Private Sub Txtcopy_Change()
    '/******************************************************************************
    If Txtcopy.Text = "" Then
        cmdCopy.Enabled = False
    Else

        cmdCopy.Enabled = True
    End If
End Sub
'/******************************************************************************
Sub drawfocusColour(x, y)
    Dim lprect As RECT, offsetx, offsety, char$
    Picture1.Line (x * sizeX + 1, y * sizeY + 1)-(x * sizeX + (sizeX - 1), _
            y * sizeY + (sizeY - 1)), vbHighlight, BF
    ''Picture1.FillColor = vbHighlight
    'Rectangle Picture1.hdc, x * sizeX + 1, y * sizeY + 1, x * sizeX + (sizeX), y * sizeY + (sizeY)
    ''Picture1.FillColor = vbWhite
    'Picture1.CurrentX = (x * sizeX) + 3
    'Picture1.CurrentY = (y * sizeY)
    '
    char$ = Chr$((y * 32) + (x + 1) + 31)
    offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
    offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
    Picture1.CurrentX = (x * sizeX) + offsetx
    Picture1.CurrentY = (y * sizeY) + offsety

    '
    Picture1.ForeColor = vbWhite
    'Picture1.Print Chr$((y * 32) + (x + 1) + 31);
    Picture1.Print char$;
    Picture1.ForeColor = vbBlack
    ''previousX = x
    ''previousY = y
    lprect.left = x * sizeX + 1
    lprect.top = y * sizeY + 1
    lprect.right = x * sizeX + (sizeX - 1) + 1 '- 1
    lprect.bottom = y * sizeY + (sizeY - 1) + 1  '- 1

    DrawFocusRect Picture1.hdc, lprect
End Sub

'/******************************************************************************
Private Sub Txtcopy_GotFocus()
    '/******************************************************************************
    '    ''Debug.Print "1   tgoptfcs"

    '        'MsgBox "gfocus"
    '        Txtcopy.SelStart = 0
    '        Txtcopy.SelLength = Len(Txtcopy.Text)
    '    End If

End Sub

'Private Sub Txtcopy_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
'        MsgBox "tkdwn"
'        Txtcopy.SelStart = 0
'        Txtcopy.SelLength = Len(Txtcopy.Text)
'    End If
'End Sub

Private Sub Txtcopy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '

End Sub

