VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFontData 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AdressOf Font Lists Demo"
   ClientHeight    =   12510
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   17415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12510
   ScaleWidth      =   17415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Speed Test"
      Height          =   2775
      Left            =   13200
      TabIndex        =   50
      Top             =   9630
      Width           =   4095
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   3855
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
         Begin VB.CommandButton Command1 
            Caption         =   "Speed Test"
            Height          =   495
            Left            =   0
            TabIndex        =   51
            Top             =   1920
            Width           =   3855
         End
         Begin VB.ListBox FList 
            Height          =   1230
            Index           =   19
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1230
            Index           =   18
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Screen.Fonts"
            Height          =   495
            Index           =   18
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AddressOf Fonts"
            Height          =   495
            Index           =   19
            Left            =   1920
            TabIndex        =   9
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Example && RTF Code"
      Height          =   5655
      Left            =   13200
      TabIndex        =   45
      Top             =   3960
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Text            =   "FrmFontData.frx":0000
         Top             =   4320
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3000
         TabIndex        =   47
         Text            =   "10"
         Top             =   240
         Width           =   705
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3720
         TabIndex        =   46
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "Text2"
         BuddyDispid     =   196616
         OrigLeft        =   16200
         OrigTop         =   240
         OrigRight       =   16455
         OrigBottom      =   615
         Max             =   100
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3495
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6165
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"FrmFontData.frx":0006
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Font Size"
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "All Fonts"
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   17175
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   16935
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   360
         Width           =   16935
         Begin VB.ListBox FList 
            Height          =   2985
            Index           =   17
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   16935
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   16935
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Font Type"
      Height          =   4095
      Left            =   120
      TabIndex        =   35
      Top             =   8310
      Width           =   6255
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   6015
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   6015
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   21
            Left            =   4080
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   20
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   16
            Left            =   4125
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   2295
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   15
            Left            =   2070
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   2295
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   14
            Left            =   15
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   2295
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   13
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   375
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   21
            Left            =   4080
            TabIndex        =   6
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   20
            Left            =   2040
            TabIndex        =   7
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   16
            Left            =   4080
            TabIndex        =   11
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   15
            Left            =   2040
            TabIndex        =   12
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   14
            Left            =   15
            TabIndex        =   13
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   13
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CharacterSets (OtherCharSets show their character set in brackets) "
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   6135
      Width           =   12495
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   12255
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   12255
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   12
            Left            =   8220
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   3975
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   6
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   9
            Left            =   2055
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   10
            Left            =   4110
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   11
            Left            =   6165
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   12
            Left            =   8220
            TabIndex        =   15
            Top             =   0
            Width           =   3975
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   9
            Left            =   2055
            TabIndex        =   0
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   10
            Left            =   4110
            TabIndex        =   17
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   11
            Left            =   6165
            TabIndex        =   16
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fixed | Proportional "
      Height          =   2175
      Left            =   8400
      TabIndex        =   42
      Top             =   8310
      Width           =   4215
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   3975
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   7
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   8
            Left            =   2055
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   8
            Left            =   2055
            TabIndex        =   1
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Family"
      Height          =   2175
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   12495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   12255
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   12255
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   0
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   1
            Left            =   2055
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   2
            Left            =   4110
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   3
            Left            =   6165
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   4
            Left            =   8220
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox FList 
            Height          =   1425
            Index           =   5
            Left            =   10275
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   1
            Left            =   2055
            TabIndex        =   18
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   2
            Left            =   4110
            TabIndex        =   8
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   3
            Left            =   6165
            TabIndex        =   5
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   4
            Left            =   8220
            TabIndex        =   4
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label FLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   375
            Index           =   5
            Left            =   10275
            TabIndex        =   3
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDetails 
      Caption         =   "Details"
      Begin VB.Menu mnuDetailsOpt 
         Caption         =   "Standard"
         Index           =   0
      End
      Begin VB.Menu mnuDetailsOpt 
         Caption         =   "NEWTEXTMETRICS"
         Index           =   1
      End
      Begin VB.Menu mnuDetailsOpt 
         Caption         =   "LOGFONT"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "FrmFontData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Command1_Click()

  Dim i As Long, Res As String
  Dim Tim1 As Single, Tim2 As Single

    Command1.Enabled = False
    Command1.Caption = "Testing..."
    FList(19).Clear
    FList(18).Clear

    FLabel(18).Caption = "Screen.Fonts : testing"
    FLabel(18).Refresh
    Tim1 = Timer
    For i = 0 To Screen.FontCount - 1
        FList(18).AddItem Screen.Fonts(i)
    Next i
    Tim1 = Timer - Tim1
    FLabel(18).Caption = Screen.FontCount - 1 & " Screen.Fonts : " & Tim1
    MsgBox "That was the slow way...", , "SpeedTest"
    FLabel(19).Caption = "AddressOf Fonts: testing"
    FLabel(19).Refresh
    DoEvents
    Tim2 = Timer
    FillListWithFontData FList(19), All
    Tim2 = Timer - Tim2
    FLabel(19).Caption = FList(19).ListCount & " AddressOf Fonts:" & Tim2

    If Tim2 < Tim1 Then
        Res = "AdressOf faster by " & Tim1 / Tim2 * 100 & "%"
      Else 'This should never occur but is here just in case'NOT TIM2...
        Res = "Screen Fonts faster by " & Tim2 - Tim1
    End If
    MsgBox "And that was the fast way..." & vbCr & Res, , "SpeedTest"
    Frame7.Caption = "Speed Test : " & Res
    Command1.Caption = "Speed Test"
    Command1.Enabled = True

End Sub

Private Sub DoFont(L As ListBox)

  Dim fn As String

    With RichTextBox1
        .SelStart = 0
        .SelLength = Len(.Text)
        fn$ = Trim$(L.List(L.ListIndex))
        If InStr(fn, vbTab) Then
            fn = Trim$(Left$(fn, InStr(fn, vbTab) - 1))
        End If
        .SelFontName = fn
        .SelFontSize = val(Text2.Text)
    End With 'RICHTEXTBOX1
    'RichTextBox1_SelChange

End Sub

Private Sub FList_Click(Index As Integer)

    DoFont FList(Index)

End Sub

Private Sub FList_LostFocus(Index As Integer)

    FList(Index).ListIndex = -1

End Sub

Private Sub Form_Initialize()

  '::) Ulli's VB Code Formatter V2.13.6 put this here but he is too modest
  'to leave the message here when you rerun the Formatter.
  'His program is a great addition to any coders utilities
  'DOWNLOAD IT NOW AT www.planet-source-code.com!!!
  'If you create a manifest file and are using XP then compiled program will
  'use XP style for all controls not just form caption bars

    InitCommonControls

End Sub

Private Sub Form_Load()

    Text2.Text = RichTextBox1.Font.Size
    mnuDetailsOpt_Click 0
    FList(17).ListIndex = 0
    DoFont FList(17)
    FillListWithFontData FList(0), DontCare
    FLabel(0).Caption = "DontCare(None) : " & FList(0).ListCount

    FillListWithFontData FList(1), Roman
    FLabel(1).Caption = "Roman : " & FList(1).ListCount

    FillListWithFontData FList(2), Swiss
    FLabel(2).Caption = "Swiss : " & FList(2).ListCount

    FillListWithFontData FList(3), Modern
    FLabel(3).Caption = "Modern : " & FList(3).ListCount
    FillListWithFontData FList(4), Script
    FLabel(4).Caption = "Script : " & FList(4).ListCount
    FillListWithFontData FList(5), Decorative
    FLabel(5).Caption = "Decorative : " & FList(5).ListCount
    Frame1.Caption = "FontFamily: " & FList(0).ListCount + FList(1).ListCount + FList(2).ListCount + FList(3).ListCount + FList(4).ListCount + FList(5).ListCount

    FillListWithFontData FList(7), Fixed
    FLabel(7).Caption = "Fixed : " & FList(7).ListCount
    FillListWithFontData FList(8), Propertional
    FLabel(8).Caption = "Propertional : " & FList(8).ListCount

    Frame2.Caption = "Fixed | Proportional  : " & FList(7).ListCount + FList(8).ListCount

    FillListWithFontData FList(6), Symbol
    FLabel(6).Caption = "Symbol : " & FList(6).ListCount
    FillListWithFontData FList(9), Ansi
    FLabel(9).Caption = "Ansi : " & FList(9).ListCount
    FillListWithFontData FList(10), Oem
    FLabel(10).Caption = "Oem : " & FList(10).ListCount

    FillListWithFontData FList(11), Default
    FLabel(11).Caption = "Default : " & FList(11).ListCount

    FillListWithFontData FList(12), OtherCharsLong
    FLabel(12).Caption = "OtherCharSets : " & FList(12).ListCount
    Frame3.Caption = "CharacterSets (OtherCharSets show their character set in brackets) : " & FList(6).ListCount + FList(9).ListCount + FList(10).ListCount + FList(11).ListCount + FList(12).ListCount

    FillListWithFontData FList(13), TrueType
    FLabel(13).Caption = "TrueType : " & FList(13).ListCount
    FillListWithFontData FList(14), Vector
    FLabel(14).Caption = "Vector : " & FList(14).ListCount
    FillListWithFontData FList(15), Raster
    FLabel(15).Caption = "Raster : " & FList(15).ListCount

    FillListWithFontData FList(16), Device
    FLabel(16).Caption = "Device : " & FList(16).ListCount

    FillListWithFontData FList(20), OpenTypeTT
    FLabel(20).Caption = "OpenType TT : " & FList(20).ListCount
    FillListWithFontData FList(21), OpenTypePS
    FLabel(21).Caption = "OpenType PS : " & FList(21).ListCount

    Frame4.Caption = "Font Type : " & FList(13).ListCount + FList(14).ListCount + FList(15).ListCount + FList(16).ListCount + FList(20).ListCount + FList(21).ListCount

End Sub

Private Sub mnuAbout_Click()

    About

End Sub

Private Sub mnuDetailsOpt_Click(Index As Integer)

  Dim OffSet As Integer, CurFont As Long

    'If your system's fonts do not line up properly with the headings
    'adjust the value of OffSet below.
    'The other headings are code based and should line up automatically
    'when you do this. This assumes that you have 'MS San Serif' if you don't for some reason then
    'select another fixed font for the Label and ListBox
    OffSet = 49
    CurFont = FList(17).ListIndex
    Select Case Index
      Case 0 'Standard
        Label1.Caption = "Name" & Space$(OffSet) & _
                         "Pitch" & Space$(17) & _
                         "Family" & Space$(9) & _
                         "CharSet" & Space$(17) & _
                         "Type" & Space$(20) & _
                         "NTM (PF-----------CS-------------OR---------AND)" & Space$(11) & _
                         "LF (PF-----------CS------------OR----------AND)    NTMFlags"

        FillListWithFontData FList(17), LongData
        Frame5.Caption = "All Fonts : " & FList(17).ListCount
      Case 1
        'NewTxtMetrics
        Label1.Caption = "Name" & Space$(OffSet) & _
                         "Height Ascent Des'nt InLead ExLead AvgWid MaxWid Weight OvrHng DAspcX DAspcY 1stChr LstChr DefChar BrkChr Italic Underl StrkOt PtcFam ChrSet     Flags          SizeEM   CellHg   AvgWid"
        FillListWithFontData FList(17), NTxtMetrics
        Frame5.Caption = "All Fonts : " & FList(17).ListCount & " NewTextMetrics"
      Case 2
        'LogFont
        Label1.Caption = "Name" & Space$(OffSet) & _
                         "Height Width  Escap Orient Weight Italic Under StrkOut CharSet OutPrec ClipPrec Quality PandF    Facename"
        FillListWithFontData FList(17), LgFont
        Frame5.Caption = "All Fonts : " & FList(17).ListCount & " LogFont"
    End Select
    FList(17).ListIndex = CurFont
Frame5.Caption = Frame5.Caption & " (Does not include secondary Bold, Italic etc fontname members where they exist)"
End Sub

Private Sub mnufileExit_Click()

    End

End Sub

Private Sub RichTextBox1_SelChange()

    Text1.Text = RichTextBox1.TextRTF

End Sub

Private Sub Text2_Change()

    With RichTextBox1
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = val(Text2.Text)
    End With 'RICHTEXTBOX1

End Sub

':) Ulli's VB Code Formatter V2.13.6 (11/09/2002 10:33:22 AM) 3 + 214 = 217 Lines
