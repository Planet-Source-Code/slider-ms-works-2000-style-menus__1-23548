VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInfoTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cInfoTip
' Author:       Slider
' Date:         07/06/2001
' Version:      01.00.00
' Description:  Self-contained multiline ToolTip class
' Edit History: 01.00.00 07/06/01 Initial Release
'
'===========================================================================

Option Explicit

Dim hWndTip As Long

Private mhWndParent As Long
Private msCaption As String

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_EXPANDTABS = &H40
Private Const LF_FACESIZE = 32
Private Const WS_POPUP = &H80000000
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const BDR_RAISEDOUTER = &H1
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const HWND_TOPMOST = -1


Private Type POINTAPI   ' pt
  X As Long
  Y As Long
End Type

Private Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

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
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSmsCaptionWidth As Long
        iSmsCaptionHeight As Long
        lfSmsCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type


Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, rc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Property Let Caption(txt As String)
    
    msCaption = txt
    
End Property

Property Get Caption() As String
    
    Caption = msCaption
    
End Property

Property Let hWnd(hWndParent As Long)

    mhWndParent = hWndParent
    
End Property

Public Sub Show(Optional TipText As String = "")

    Const margin = 3

    Dim p        As POINTAPI
    Dim rc       As RECT
    Dim ncm      As NONCLIENTMETRICS

    Dim dc       As Long
    Dim hbrush   As Long
    Dim rgb      As Long
    Dim FontName As String
    Dim fntTip   As New StdFont
    Dim fnt      As IFont

    Dim lLoop    As Long
    Dim sTmp     As String
    Dim numRows  As Long
    Dim maxLen   As Long
    Dim maxTxt   As String
    Dim arrTxt() As String

    Dim X        As Long
    Dim Y        As Long
    Dim dX       As Long
    Dim dY       As Long

    If Len(TipText) = 0 Then
        sTmp = msCaption
    Else
        sTmp = TipText
    End If

    '## Cleaning
    sTmp = Replace(sTmp, vbCrLf, vbCr)
    If Right$(sTmp, 1) = vbCr Then
        sTmp = Left$(sTmp, Len(sTmp) - 1)
    End If

    '## Maximum row length
    arrTxt = Split(sTmp, vbCr)
    sTmp = ""
    numRows = UBound(arrTxt)
    For lLoop = 0 To numRows
        sTmp = sTmp & arrTxt(lLoop) & vbCrLf
        If maxLen < Len(arrTxt(lLoop)) Then
            maxTxt = arrTxt(lLoop) & Chr$(0)
            maxLen = Len(arrTxt(lLoop))
        End If
    Next
    Erase arrTxt
    sTmp = Left$(sTmp, Len(sTmp) - 2) '## Last vbCrLf

    '## Tooltip Font Information
    ncm.cbSize = Len(ncm)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(ncm), ncm, 0

    With ncm.lfStatusFont
        For lLoop = 1 To LF_FACESIZE
            If .lfFaceName(lLoop) = 0 Then
                Exit For
            End If
            FontName = FontName & Chr$(.lfFaceName(lLoop))
        Next

        fntTip.Charset = .lfCharSet
        fntTip.Italic = .lfItalic
        fntTip.Name = FontName
        fntTip.Size = -.lfHeight * 72 / GetDeviceCaps(GetDC(0), LOGPIXELSY)
        fntTip.Strikethrough = .lfStrikeOut
        fntTip.Underline = .lfUnderline
        fntTip.Weight = .lfWeight
    End With

    '## Create Window
    If hWndTip = 0 Then
        hWndTip = CreateWindowEx(0, "STATIC", _
            vbNullString, WS_POPUP, _
            0, 0, _
            0, 0, _
            mhWndParent, 0, _
            App.hInstance, 0)
    End If

    dc = GetDC(hWndTip)

    Set fnt = fntTip
    SelectObject dc, fnt.hFont

    '## Calculate Size
    DrawText dc, sTmp, Len(sTmp), rc, DT_TOP Or DT_LEFT Or DT_CALCRECT Or DT_EXPANDTABS
    With rc
        .Bottom = .Bottom + 2 * margin
        .Right = .Right + 2 * margin
    End With

    dX = rc.Right - rc.Left
    dY = rc.Bottom - rc.Top

    '## Get mouse position
    GetCursorPos p
    X = p.X
    Y = p.Y + 22

    '## Check Position versus Screen size
    With Screen
        If X + dX > .Width / .TwipsPerPixelX Then
            X = .Width / .TwipsPerPixelX - dX - margin
        End If
        
        If Y + dY > .Height / .TwipsPerPixelY Then
            Y = p.Y - margin - dY
        End If
    End With

    '## Show Window
    SetWindowPos hWndTip, HWND_TOPMOST, X, Y, dX, dY, SWP_SHOWWINDOW Or SWP_NOACTIVATE
    DoEvents

    '## Fill with Tooltip Backcolor
    OleTranslateColor vbInfoBackground, 0, rgb
    hbrush = CreateSolidBrush(rgb)
    FillRect dc, rc, hbrush
    DeleteObject hbrush

    '## Draw 3D border
    DrawEdge dc, rc, BDR_RAISEDOUTER, BF_RECT

    '## Set text colors
    OleTranslateColor vbInfoBackground, 0, rgb
    SetBkColor dc, rgb

    OleTranslateColor vbInfoText, 0, rgb
    SetTextColor dc, rgb

    '## Adjust Text rectangle
    With rc
        .Left = .Left + margin
        .Top = .Top + margin
        .Bottom = .Bottom - margin
        .Right = .Right - margin
    End With

    '## Paint text
    DrawText dc, sTmp, Len(sTmp), rc, DT_TOP Or DT_LEFT Or DT_EXPANDTABS

End Sub

Public Sub Hide()
    If hWndTip <> 0 Then
        DestroyWindow hWndTip
    End If
    hWndTip = 0
End Sub

Private Sub Class_Terminate()
    Hide
End Sub
