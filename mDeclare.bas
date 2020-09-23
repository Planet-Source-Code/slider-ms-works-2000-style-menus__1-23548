Attribute VB_Name = "modDelcares"
'===========================================================================
'
' Module Name:  modDeclares
' Author:       Slider
' Date:         29/05/01
' Version:      01.00.00
' Description:  API Declarations and reusable code units
' Edit History: 01.00.00 29/05/01 Initial Release
'
'===========================================================================

Option Explicit

Public Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type

Type POINTAPI  '  8 Bytes
    X As Long
    Y As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Declare Function winSetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function winGetFocus Lib "user32" Alias "GetFocus" () As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Focus and activation functions
Public Const GWL_STYLE = (-16)
Public Const WS_CHILD = &H40000000

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

' Rectangle functions:
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbrush As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' Background Modes
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2
Public Const BKMODE_LAST = 2

' Pen Styles
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const PS_USERSTYLE = 7
Public Const PS_ALTERNATE = 8
Public Const PS_STYLE_MASK = &HF

Public Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign& Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long)
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Public Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function GetTextExtentPoint32& Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size)

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long

Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Const LF_FACESIZE = 32

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

Public Type LOGFONT
   lfHeight As Long ' The font size (see below)
   lfWidth As Long ' Normally you don't set this, just let Windows create the Default
   lfEscapement As Long ' The angle, in 0.1 degrees, of the font
   lfOrientation As Long ' Leave as default
   lfWeight As Long ' Bold, Extra Bold, Normal etc
   lfItalic As Byte ' As it says
   lfUnderline As Byte ' As it says
   lfStrikeOut As Byte ' As it says
   lfCharSet As Byte ' As it says
   lfOutPrecision As Byte ' Leave for default
   lfClipPrecision As Byte ' Leave for default
   lfQuality As Byte ' Leave for default
   lfPitchAndFamily As Byte ' Leave for default
   lfFaceName(LF_FACESIZE) As Byte ' The font name converted to a byte array
End Type

Type LOGBRUSH     '  12 Bytes
     lbStyle As Long
     lbColor As Long
     lbHatch As Long
End Type

Public Type DRAWTEXTPARAMS
    cbSize        As Long
    iTabLength    As Long
    iLeftMargin   As Long
    iRightMargin  As Long
    uiLengthDrawn As Long
End Type

Public Type TEXTMETRIC   '  53 Bytes
        tmHeight           As Long
        tmAscent           As Long
        tmDescent          As Long
        tmInternalLeading  As Long
        tmExternalLeading  As Long
        tmAveCharWidth     As Long
        tmMaxCharWidth     As Long
        tmWeight           As Long
        tmOverhang         As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar        As Byte
        tmLastChar         As Byte
        tmDefaultChar      As Byte
        tmBreakChar        As Byte
        tmItalic           As Byte
        tmUnderlined       As Byte
        tmStruckOut        As Byte
        tmPitchAndFamily   As Byte
        tmCharSet          As Byte
End Type

Public Enum eTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
                   
Private Const WM_SETREDRAW = &HB

Public Function LockControl(objX As Object, cLock As Boolean)

   If cLock Then
      ' Disable the Redraw flag for the specified window
      Call SendMessage(objX.hWnd, WM_SETREDRAW, False, 0)
   Else
      ' Enable the Redraw flag for the specified window, and repaint
      Call SendMessage(objX.hWnd, WM_SETREDRAW, True, 0)
      objX.Refresh
   End If

End Function

Public Sub CenterForm(Frm As Form)

    Dim ClientRect  As RECT
    Dim TaskBarRect As RECT
    Dim X           As Variant
    Dim Y           As Variant
    Dim lRetVal     As Long

    If Frm.MDIChild Then                                    '## Check if the form is a MDIChild.
        GetClientRect GetParent(Frm.hWnd), ClientRect       '## Center it in the MDIParent.
    Else
        Call GetClientRect(GetDesktopWindow(), ClientRect)  '## Get the Desktop area
        lRetVal = FindWindow("Shell_TrayWnd", vbNullString) '## Check for the Task Bar.
        If lRetVal Then                                     '## If there is a taskbar, then adjust the ClientRect.
            Call GetWindowRect(lRetVal, TaskBarRect)
            If (TaskBarRect.Right - TaskBarRect.Left) > (TaskBarRect.Bottom - TaskBarRect.Top) Then
                If TaskBarRect.Top <= 0 Then                '## TaskBar at the Top of Screen.
                    ClientRect.Top = ClientRect.Top + TaskBarRect.Bottom
                Else                                        '## TaskBar at the Bottom of Screen.
                    ClientRect.Bottom = ClientRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top)
                End If
            Else
                If TaskBarRect.Left <= 0 Then               '## TaskBar is on the Left side of the Screen.
                    ClientRect.Left = ClientRect.Left + TaskBarRect.Right
                Else                                        '## TaskBar is on the Right side of the Screen.
                    ClientRect.Right = ClientRect.Right - (TaskBarRect.Right - TaskBarRect.Left)
                End If
            End If
        End If
    End If
    With Frm                                                '## Center the Form
        X = (((ClientRect.Right - ClientRect.Left) * Screen.TwipsPerPixelX) - .Width) \ 2
        Y = (((ClientRect.Bottom - ClientRect.Top) * Screen.TwipsPerPixelY) - .Height) \ 2
        .Move X, Y
    End With

End Sub

Public Sub HiLite(txtBox As TextBox)
    With txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Public Sub DrawText(TextLine As String, _
                    hdc As Long, _
                    DrawRect As RECT, _
                    Flags As eTextFlags, _
           Optional LeftMargin As Long = 0, _
           Optional RightMargin As Long = 0, _
           Optional Kerning As Long = 0)

    Dim DrawParams As DRAWTEXTPARAMS

    With DrawParams
        .cbSize = Len(DrawParams)
        .iLeftMargin = LeftMargin
        .iRightMargin = RightMargin
    End With
    SetTextCharacterExtra hdc, Val(Kerning)
    DrawTextEx hdc, TextLine, Len(TextLine), DrawRect, Flags, DrawParams

End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Sub OLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)

    Dim sFont As String
    Dim iChar As Integer

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
       sFont = fntThis.Name
       ' There is a quicker way involving StrConv and CopyMemory, but
       ' this is simpler!:
       For iChar = 1 To Len(sFont)
           .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
       Next iChar
       ' Based on the Win32SDK documentation:
       .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
       .lfItalic = fntThis.Italic
       If (fntThis.Bold) Then
           .lfWeight = FW_BOLD
       Else
           .lfWeight = FW_NORMAL
       End If
       .lfUnderline = fntThis.Underline
       .lfStrikeOut = fntThis.Strikethrough
       .lfCharSet = fntThis.Charset
   End With

End Sub
