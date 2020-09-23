Attribute VB_Name = "FontData"

Option Explicit
'MODIFEID FROM 'AddressOf Operator Example' ARTICLE IN MSDN LIBRARY VISUAL STUDIO 6

'Font enumeration types
Public Const LF_FACESIZE  As Long = 32
Public Const LF_FULLFACESIZE As Long = 64

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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
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

Public Const NTM_BOLD As Long = &H20&
Public Const NTM_ITALIC As Long = &H1&
Private Const NTM_PS_OPENTYPE As Long = &H20000
Private Const NTM_TT_OPENTYPE As Long = &H40000
Private Const NTM_REGULAR As Long = &H40&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH  As Long = &H1
Public Const TMPF_VECTOR As Long = &H2
Public Const TMPF_DEVICE  As Long = &H8
Public Const TMPF_TRUETYPE  As Long = &H4
Public Const OEM_FIXED_FONT As Long = 10

'  EnumFonts Masks

Private Const FF_DONTCARE As Long = 0
Private Const FF_ROMAN  As Long = 16
Private Const FF_SWISS As Long = 32
Private Const FF_MODERN  As Long = 48
Private Const FF_SCRIPT As Long = 64
Private Const FF_DECORATIVE As Long = 80

'Character Sets
Public Const ANSI_CHARSET As Long = 0
Public Const DEFAULT_CHARSET As Long = 1
Public Const SYMBOL_CHARSET As Long = 2
Public Const SHIFTJIS_CHARSET As Long = 128
Public Const HANGEUL_CHARSET As Long = 129
Public Const HANGUL_CHARSET As Long = 129
Public Const CHINESEBIG5_CHARSET As Long = 136
Public Const OEM_CHARSET As Long = 255
Public Const JOHAB_CHARSET As Long = 130
Public Const HEBREW_CHARSET As Long = 177
Public Const ARABIC_CHARSET As Long = 178
Public Const GREEK_CHARSET As Long = 161
Public Const TURKISH_CHARSET As Long = 162
Public Const THAI_CHARSET  As Long = 222
Public Const EASTEUROPE_CHARSET As Long = 238
Public Const RUSSIAN_CHARSET  As Long = 204
Public Const MAC_CHARSET As Long = 77
Public Const BALTIC_CHARSET As Long = 186
Public Const VIETNAMESE_CHARSET As Long = 163
Public Const GB2312_CHARSET As Long = 134

Private Const WM_USER As Long = &H400

Private Const LB_SETTABSTOPS As Long = &H192        ' Has changed in Win32.
Private Const LB_SETHORIZONTALEXTENT As Long = &H194   ' Has changed in Win32.
Private Declare Function SendMessage Lib "user32" Alias _
                          "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                          ByVal wParam As Long, lParam As Long) As Long
Private scrollbarwidth   As Long    ' Width of horizontal scrollbar.
Private numtabs As Long            ' Number of tabs needed.
Private tabstops() As Long         ' Array of value of tab stop of columnn

'My variables
Private FontCol As New Collection 'Use collection because you dont have to worry about dimensioning it
Private Const MySep As String = vbTab

Public Enum FontTypes
    All
    DontCare
    Roman
    Swiss
    Modern
    Script
    Decorative
    Fixed
    Propertional
    Symbol
    Ansi
    Oem
    Default
    OtherChars
    OtherCharsLong
    TrueType
    OpenTypeTT
    OpenTypePS
    Raster
    Vector
    Device
    LongData
    NTxtMetrics
    LgFont
End Enum
Private LMode As FontTypes
Rem Mark Off
''stops Ulli 's Code Formatter from noticing these as Duplicated Name without Scope or TypeCasting
#If False Then      'Enforce Case for Enums (does not compile)
Dim All             'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
im DontCare
Dim Roman
Dim Swiss
Dim Modern
Dim Script
Dim Decorative
Dim Fixed
Dim Propertional
Dim Symbol
Dim Ansi
Dim Oem
Dim Default
Dim OtherChars
Dim OtherCharsLong
Dim WithData
Dim TrueType
Dim OpenTypeTT
Dim OpenTypePS
Dim Raster
Dim Vector
Dim Device
Dim LongData
Dim NTxtMetrics
Dim LgFont
#End If
Rem Mark On

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Sub About()
'Copyright 2002 Roger Gilchrist
  Dim msg As String

    msg = "This is a spin off from my project 'clsExtendedRTF'. " & vbCr & _
          "I was looking for access to font family and character set data." & vbCr & _
          "I had hoped to incorperate it in the class but AddressOf doesn't work in classes." & vbCr & _
          "So here it is a a bas file." & vbCr & vbCr & _
          "The module uses two versions of font data from two different Types." & vbCr & _
          "LogFont (LF) and NewTextMetrics (NTM)." & vbCr & _
          "PF is the .lfPitchAndFamily or .tmPitchAndFamily values returned by the Types." & vbCr & _
          "CS is the .lfCharSet or .tmCharSet values returned by the Types." & vbCr & _
          "ntmFlags holds the OpenType indicator." & vbCr & _
          "I have ORed and ANDed both of the PF's with CS's." & vbCr & _
          "The Demo uses NewTextMetrics to fill the lists except for FontName from LogFont." & vbCr & _
          "If the 'All Fonts' label headings do not line up adjust OffSet value in mnuDetailsOpt_Click" & vbCr & _
          "(depends on longest fontname on your system)." & vbCr & _
          "Details Menu shows 'Standard', LogFont & NewTextMetrics views of Font Data." & vbCr & _
          "Click Speed Test to compare with Screen.Fonts method (approx 1000 X faster)" & vbCr & _
           "Copyright 2002 Roger Gilchrist"

    MsgBox msg, , "AdressOf Font Lists Demo"

End Sub

Private Function BitRead(Byt As Byte, Bit As Integer) As Integer

  ' Return the value of the 2 to the nth power bit:

  Dim Mask As Long

    Mask = 2 ^ Bit ' Create a bitmask with the 2 to the nth power bit set:
    BitRead = Byt And Mask

End Function

Private Function BitTest(Byt As Byte, Bit As Integer) As Boolean

  ' The BitTest function will return True or False depending on
  ' the value of the nth bit (Bit%) of an integer (Byte%).

  Dim Mask As Long

    Mask = 2 ^ Bit
    BitTest = ((Byt And Mask) > 0)

End Function

Private Function CharSetStr(val As Integer) As String

  'Copyright 2002 Roger Gilchrist

    Select Case val
      Case ANSI_CHARSET
        CharSetStr = "ANSI"
      Case DEFAULT_CHARSET
        CharSetStr = "DEFAULT"
      Case SYMBOL_CHARSET
        CharSetStr = "SYMBOL"
      Case SHIFTJIS_CHARSET
        CharSetStr = "SHIFTJIS"
      Case HANGUL_CHARSET, HANGEUL_CHARSET ' historic misspell in WinAPI
        CharSetStr = "HANGUL"
      Case GB2312_CHARSET
        CharSetStr = ""
      Case CHINESEBIG5_CHARSET
        CharSetStr = "CHINESEBIG5"
      Case OEM_CHARSET
        CharSetStr = "OEM"
      Case JOHAB_CHARSET
        CharSetStr = "JOHAB"
      Case HEBREW_CHARSET
        CharSetStr = "HEBREW"
      Case ARABIC_CHARSET
        CharSetStr = "ARABIC"
      Case GREEK_CHARSET
        CharSetStr = "GREEK"
      Case TURKISH_CHARSET
        CharSetStr = "TURKISH"
      Case VIETNAMESE_CHARSET
        CharSetStr = ""
      Case THAI_CHARSET
        CharSetStr = "THAI"
      Case EASTEUROPE_CHARSET
        CharSetStr = "EASTEUROPE"
      Case RUSSIAN_CHARSET
        CharSetStr = "RUSSIAN"
      Case MAC_CHARSET
        CharSetStr = "MAC"
      Case BALTIC_CHARSET
        CharSetStr = "BALTIC"
      Case Else
        CharSetStr = "Unknown"
    End Select

End Function

Private Function DoIt(lpNTM As NEWTEXTMETRIC, lpNLF As LOGFONT, Attach As String) As Boolean

  'Copyright 2002 Roger Gilchrist
  'test if the font is suitable for the requested type
  'makes attachments if necessary

  Dim Family As Long
  Dim Pitch As Boolean
  Dim ChrSet As Integer

    Family = &HF0 And lpNTM.tmPitchAndFamily
    Pitch = BitTest(lpNTM.tmPitchAndFamily, 0)
    ChrSet = CInt(lpNTM.tmCharSet)
    Select Case LMode
      Case All
        DoIt = True
      Case DontCare
        DoIt = FF_DONTCARE = Family
      Case Decorative
        DoIt = FF_DECORATIVE = Family
      Case Modern
        DoIt = FF_MODERN = Family
      Case Script
        DoIt = FF_SCRIPT = Family
      Case Roman
        DoIt = FF_ROMAN = Family
      Case Swiss
        DoIt = FF_SWISS = Family
      Case Fixed
        DoIt = Not Pitch
      Case Propertional
        DoIt = Pitch
      Case Symbol
        DoIt = ChrSet = SYMBOL_CHARSET
      Case Ansi
        DoIt = ChrSet = ANSI_CHARSET
      Case Oem
        DoIt = ChrSet = OEM_CHARSET
      Case Default
        DoIt = ChrSet = DEFAULT_CHARSET
      Case OtherChars, OtherCharsLong
        DoIt = ChrSet <> OEM_CHARSET And ChrSet <> ANSI_CHARSET And _
               ChrSet <> SYMBOL_CHARSET And ChrSet <> DEFAULT_CHARSET
        If LMode = OtherCharsLong Then
            Attach = MySep & CharSetStr(ChrSet)
        End If
      Case Raster

        DoIt = FontType(lpNTM) = "Raster"
      Case Vector
        DoIt = FontType(lpNTM) = "Vector"
      Case TrueType
        DoIt = (FontType(lpNTM) = "TrueType")
      Case Device
        DoIt = (lpNTM.tmPitchAndFamily And TMPF_DEVICE)
      Case LongData
        DoIt = True
        Attach = MySep & LongDataStr(lpNTM, lpNLF)
      Case NTxtMetrics
        DoIt = True
        Attach = MySep & NewTextMetricsStr(lpNTM)
      Case LgFont
        DoIt = True
        Attach = MySep & LogFontStr(lpNLF)
      Case OpenTypeTT
        DoIt = FontType(lpNTM) = "OpenTypeTT" '.ntmFlags And NTM_TT_OPENTYPE
      Case OpenTypePS
        DoIt = FontType(lpNTM) = "OpenTypePS" '.ntmFlags And NTM_TT_OPENTYPE
        'NTM_PS_OPENTYPE
        'NTM_TT_OPENTYPE
      Case Else
    End Select

End Function

Private Function EnumFontFamProcToCollection(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, Coll As Collection) As Long

  'MODIFEID FROM 'AddressOf Operator Example' ARTICLE IN MSDN LIBRARY VISUAL STUDIO 6
'Modification Copyright 2002 Roger Gilchrist
  Dim FaceName As String
  Dim Attach As String

    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    FaceName = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    If DoIt(lpNTM, lpNLF, Attach) Then
        Coll.Add FaceName & Attach
    End If
    EnumFontFamProcToCollection = 1

End Function

Private Sub Fill_List_Formated(Lb As ListBox, Coll As Collection, Sep As String, Ignore As String, Optional WhiteSpaceSize As Integer = 2)

  'MODIFEID FROM THIS ARTICLE IN MSDN LIBRARY VISUAL STUDIO 6
  'How to Fill a List Box with Snapshot When Contents are Unknown
  'Article ID: Q141026

  'MODIFICATIONS
  'Copyright 2002 Roger Gilchrist
  'Original name: Fill_List
  'Coll : the collection into which the font data was dumped
  'Sep : the separator character used in each member of the collection
  'WhiteSpaceSize: Amount of white space between columns. (Replaces Const NUMCHARS in original)
  'Ignore: if a member of the collection contains this string don't add it to the ListBox
  'Lb.Parent replaces direct calls to Form

  ' Temporary variables to preserve form font settings:

  Dim hold_fontname As String, hold_fontsize As Integer
  Dim hold_fontbold As Integer, hold_fontitalic As Integer
  Dim hold_fontstrikethru As Integer, hold_fontunderline  As Integer

  Dim whiteSpace As Integer, accumtabstops As Integer, dialogUnits As Integer
  Dim fieldVal As String, listline As String
  Dim avgWidth As Single
  Dim i As Integer                ' Used in For Next loops.
  Dim biggest_value() As Single   ' Array of longest string of columns.
  Dim retval As Long              ' Return value of SendMessage function
  Dim fieldvals As Variant

  Dim j As Integer 'MODIFICATION Used to cycle through each member of collection

    ' Save form's font settings so we can use the form to calculate the
    ' TextWidth / Height of the strings to go into the list box.
    If Ignore = "" Then 'MODIFICATION
        Ignore = "XXXXXXXXX"
    End If
    hold_fontname = Lb.Parent.FontName
    hold_fontsize = Lb.Parent.FontSize
    hold_fontbold = Lb.Parent.FontBold
    hold_fontitalic = Lb.Parent.FontItalic
    hold_fontstrikethru = Lb.Parent.FontStrikethru
    hold_fontunderline = Lb.Parent.FontUnderline

    ' Set form font settings to be identical to list box.
    Lb.Parent.FontName = Lb.FontName
    Lb.Parent.FontSize = Lb.FontSize
    Lb.Parent.FontBold = Lb.FontBold
    Lb.Parent.FontItalic = Lb.FontItalic
    Lb.Parent.FontStrikethru = Lb.FontStrikethru
    Lb.Parent.FontUnderline = Lb.FontUnderline

    ' Get the average character width of the current list box font
    ' (in pixels) using the form's TextWidth width method.
    avgWidth = Lb.Parent.TextWidth("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
    avgWidth = avgWidth / Screen.TwipsPerPixelX / 52

    ' Set the white space you want between columns.
    whiteSpace = avgWidth * WhiteSpaceSize

    ReDim biggest_value(0 To Coll.Count - 1)
    ReDim tabstops(1 To Coll.Count)
    ' Loop through the values for each member of the collection
    ' Calculate the width required for each value to fit in the list box.
    'Also, build each line of the list box and add it to the list as you go.
    For i = 1 To Coll.Count
        fieldvals = Split(Coll.Item(i), Sep)

        For j = LBound(fieldvals) To UBound(fieldvals)
            ' fieldVal = sn(i) & ""       ' Append "" in case of a null field.

            ' The LB_SETTABSTOP message requires coordinates in dialog units
            ' (roughly 4 *, the average character width in pixels).
            dialogUnits = ((Lb.Parent.TextWidth(fieldvals(j)) / Screen.TwipsPerPixelX + whiteSpace) \ avgWidth) * 4
            If dialogUnits > biggest_value(j) Then
                biggest_value(j) = dialogUnits
            End If

            listline = listline & fieldvals(j) & vbTab
        Next j
        'MODIFICATION
        'Remove last tab and biggestValue unit so that an un separated list doesn't
        'generate a column and activate the Horizontal scrollbar
        listline = Left$(listline, InStrRev(listline, vbTab) - 1)
        'dirty trick; j is always one more than highest value of the For Next structure
        biggest_value(j - 1) = 0

        If InStr(listline, Ignore$) = 0 Then ' force tabs to size of force size member of collection but don't add it
            Lb.AddItem listline
        End If
        listline = ""

    Next i

    ' Fill the tabstops() array with the position of each tab stop.
    For i = 0 To Coll.Count - 1
        accumtabstops = accumtabstops + biggest_value(i)
        tabstops(i + 1) = accumtabstops
    Next i

    ' numtabs must be a Long for Win32, Integer for Win16.
    numtabs = i
    ' Send LB_SETTABSTOP to the list box to set the position of each
    ' column.

    retval& = SendMessage(Lb.hWnd, LB_SETTABSTOPS, numtabs, tabstops(1))

    ' Set the horizontal extent just wider than the first tab stop.
    ' This produces a horizontal scroll bar on the list box.
    ' This message requires coordinates in pixels, so we convert the tab
    ' stop coordinate back from dialog units to pixels.

    scrollbarwidth = (tabstops(i) \ 4) * avgWidth
    retval& = SendMessage(Lb.hWnd, LB_SETHORIZONTALEXTENT, scrollbarwidth, 0&)

    ' Restore form's original font property settings.
    Lb.Parent.FontName = hold_fontname
    Lb.Parent.FontSize = hold_fontsize
    Lb.Parent.FontBold = hold_fontbold
    Lb.Parent.FontItalic = hold_fontitalic
    Lb.Parent.FontStrikethru = hold_fontstrikethru
    Lb.Parent.FontUnderline = hold_fontunderline

End Sub

Public Sub FillListWithFontData(Lb As ListBox, Mode As FontTypes)

  'MODIFEID FROM 'AddressOf Operator Example' ARTICLE IN MSDN LIBRARY VISUAL STUDIO 6
  'Modification Copyright 2002 Roger Gilchrist

  Dim Ignore As String
  Dim Force As String

    Lb.Clear
    Set FontCol = GetFontsAsCollection(Lb, Mode)
    Select Case LMode
      Case LongData
        'Force cols to minimum widths for the long data format
        'remember to set the Ignore string to match one of the strings of @
        Ignore = "@@@@@@@@@@@@"
        Force = "@@@@@@@@@@@@" & MySep & "@@@@@" & MySep & "@@@@@" & MySep & "@@@@@" & MySep & _
                "@@@@@@@@" & MySep & "@@@" & MySep & "@@@" & MySep & "@@@" & MySep & _
                "@@@@@" & MySep & "@@@" & MySep & "@@@" & MySep & "@@@" & MySep & "@@@@" & MySep & "@@@@@@@"
        FontCol.Add Force

      Case LgFont
        Ignore = "@@@@@@@@@@@@"
        Force = "@@@@@@@@@@@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & _
                MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & _
                MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@"
        FontCol.Add Force

      Case NTxtMetrics
        Ignore = "@@@@@@@@@@@@"
        Force = "@@@@@@@@@@@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@" & _
                MySep & "@@" & MySep & "@@@" & MySep & "@@" & MySep & "@@@" & MySep & "@@" & _
                MySep & "@@@" & MySep & "@@" & MySep & "@@@" & MySep & "@@" & MySep & "@" & _
                MySep & "@" & MySep & "@" & MySep & "@" & MySep & "@@" & MySep & "@" & _
                MySep & "@@" & MySep & "@@" & MySep & "@@" & MySep & "@@"

        FontCol.Add Force
    End Select

    If FontCol.Count Then ' don't try to manipulate listbox if there's nothing to add
        Fill_List_Formated Lb, FontCol, vbTab, Ignore, 3
    End If

End Sub

Private Function FontFamily(val As Long) As String

  'Copyright 2002 Roger Gilchrist

    Select Case val
      Case FF_DONTCARE
        FontFamily = "DontCare"
      Case FF_ROMAN
        FontFamily = "Roman"
      Case FF_SWISS
        FontFamily = "Swiss"
      Case FF_MODERN
        FontFamily = "Modern"
      Case FF_SCRIPT
        FontFamily = "Script"
      Case FF_DECORATIVE
        FontFamily = "Decor" '"Decorative"
      Case Else
        FontFamily = val
    End Select

End Function

Private Function FontType(lpNTM As NEWTEXTMETRIC) As String
'Copyright 2002 Roger Gilchrist
  Dim TPF As Byte
  Dim TFlag As Long

    TPF = lpNTM.tmPitchAndFamily
    TFlag = lpNTM.ntmFlags
    If (lpNTM.ntmFlags And NTM_TT_OPENTYPE) Then
        FontType = "OpenTypeTT"
      ElseIf (lpNTM.ntmFlags And NTM_PS_OPENTYPE) Then 'NOT (LPNTM.NTMFLAGS...
        FontType = "OpenTypePS"
      ElseIf (TPF And TMPF_TRUETYPE) Then 'NOT (LPNTM.NTMFLAGS...
        FontType = "TrueType" '1,1
      ElseIf BitTest(TPF, 1) And Not BitTest(TPF, 2) Then 'NOT (TPF...
        FontType = "Vector" '1,0
      ElseIf Not BitTest(TPF, 1) And Not BitTest(TPF, 2) Then 'NOT BITTEST(TPF,...
        FontType = "Raster" '0,0
      Else 'NOT NOT...
        'should never hit but just for safety
        FontType = "Unknown"
    End If

End Function

Public Function GetFontsAsCollection(Lb As ListBox, Mode As FontTypes) As Collection

  'MODIFEID FROM 'AddressOf Operator Example' ARTICLE IN MSDN LIBRARY VISUAL STUDIO 6
  'Even though you are not using a listbox for this EnumFontFamilies needs a hDC value
  'The Listbox does not need to have anything to do with your font stuff
  'Modification Copyright 2002 Roger Gilchrist

  Dim hDC As Long

    hDC = GetDC(Lb.hWnd) 'because EnumFontFamilies needs a Hwnd
    LMode = Mode
    Do While FontCol.Count 'clear previous if any
        FontCol.Remove 1
    Loop
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProcToCollection, FontCol
    ReleaseDC Lb.hWnd, hDC
    Set GetFontsAsCollection = FontCol

End Function

Private Function LogFontStr(lpNLF As LOGFONT) As String

  'convert LOGFONT to a string
  'Copyright 2002 Roger Gilchrist

    With lpNLF
        LogFontStr = .lfHeight ' As Long
        LogFontStr = LogFontStr & MySep & .lfWidth ' As Long
        LogFontStr = LogFontStr & MySep & .lfEscapement ' As Long
        LogFontStr = LogFontStr & MySep & .lfOrientation ' As Long
        LogFontStr = LogFontStr & MySep & .lfWeight ' As Long
        LogFontStr = LogFontStr & MySep & .lfItalic ' As Byte
        LogFontStr = LogFontStr & MySep & .lfUnderline ' As Byte
        LogFontStr = LogFontStr & MySep & .lfStrikeOut ' As Byte
        LogFontStr = LogFontStr & MySep & .lfCharSet ' As Byte
        LogFontStr = LogFontStr & MySep & .lfOutPrecision ' As Byte
        LogFontStr = LogFontStr & MySep & .lfClipPrecision ' As Byte
        LogFontStr = LogFontStr & MySep & .lfQuality ' As Byte
        LogFontStr = LogFontStr & MySep & .lfPitchAndFamily ' As Byte
        LogFontStr = LogFontStr & MySep & StrConv(lpNLF.lfFaceName, vbUnicode) ' As Byte
    End With 'LPNLF

End Function

Private Function LongDataStr(lpNTM As NEWTEXTMETRIC, lpNLF As LOGFONT) As String

  'Copyright 2002 Roger Gilchrist

  Dim OR1 As Variant, AND1 As Variant
  Dim OR2 As Variant, AND2 As Variant

    With lpNTM
        OR1 = .tmPitchAndFamily Or .tmCharSet
        AND1 = .tmPitchAndFamily And .tmCharSet
    End With 'LPNTM

    With lpNLF
        OR2 = .lfPitchAndFamily Or .lfCharSet
        AND2 = .lfPitchAndFamily And .lfCharSet
    End With 'LPNLF
LongDataStr = IIf(lpNTM.tmPitchAndFamily And TMPF_FIXED_PITCH, "Proportional", "Fixed") & MySep & _
                  FontFamily(&HF0 And lpNTM.tmPitchAndFamily) & MySep & _
                  CharSetStr(CInt(lpNTM.tmCharSet)) & MySep & _
                  FontType(lpNTM) & MySep & _
                  lpNTM.tmPitchAndFamily & MySep & _
                  lpNTM.tmCharSet & MySep & _
                  OR1 & MySep & _
                  AND1 & MySep & _
                  lpNLF.lfPitchAndFamily & MySep & _
                  lpNLF.lfCharSet & MySep & _
                  OR2 & MySep & _
                  AND2 & MySep & _
                  lpNTM.ntmFlags

End Function

Private Function NewTextMetricsStr(lpNTM As NEWTEXTMETRIC) As String

  'convert NEWTEXTMETRICS to a string
  'Copyright 2002 Roger Gilchrist

    With lpNTM
        NewTextMetricsStr = .tmHeight ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmAscent ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmDescent ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmInternalLeading ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmExternalLeading ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmAveCharWidth ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmMaxCharWidth ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmWeight ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmOverhang ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmDigitizedAspectX ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmDigitizedAspectY ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmFirstChar ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmLastChar ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmDefaultChar ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmBreakChar ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmItalic ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmUnderlined ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmStruckOut ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmPitchAndFamily ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .tmCharSet ' As Byte
        NewTextMetricsStr = NewTextMetricsStr & MySep & .ntmFlags ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .ntmSizeEM ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .ntmCellHeight ' As Long
        NewTextMetricsStr = NewTextMetricsStr & MySep & .ntmAveWidth ' As Long
    End With 'LPNTM

End Function

':) Ulli's VB Code Formatter V2.13.6 (11/09/2002 10:33:17 AM) 173 + 517 = 690 Lines
