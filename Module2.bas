Attribute VB_Name = "Module2"

Option Explicit
'Cteated By Joko Mulyono
'Email:dantex_765@hotmail.com
Public Type TYPERECT
    Left                             As Long
    Top                              As Long
    Right                            As Long
    Bottom                           As Long
End Type
Public Enum Appearance

    Flat = 0
    HalfRaised = 1
    Raised = 2
    Sunken = 3
    Etched = 4
    Bump = 5
    Line = 6
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Flat, HalfRaised, Raised, Sunken, Etched, Bump
#End If

Public Enum Warnalatar
    vbScrollBars = &H80000000       'Scroll bar color
    vbDesktop = &H80000001      'Desktop color
    vbActiveTitleBar = &H80000002  ' Color of the title bar for the active window
    vbInactiveTitleBar = &H80000003 'Color of the title bar for the inactive window
    vbMenuBar = &H80000004      'Menu background color
    vbWindowBackground = &H80000005 'Window background color
    vbWindowFrame = &H80000006      'Window frame color
    vbMenuText = &H80000007     'Color of text on menus
    vbWindowText = &H80000008       'Color of text in windows
    vbTitleBarText = &H80000009    ' Color of text in caption, size box, and scroll arrow
    vbActiveBorder = &H8000000A    ' Border color of active window
    vbInactiveBorder = &H8000000B   'Border color of inactive window
    vbApplicationWorkspace = &H8000000C ' Background color of MDI interface
    vbHighlight = &H8000000D        'Background color of items selected in a control
    vbHighlightText = &H8000000E    'Text color of items selected in a control
    vbButtonFace = &H8000000F       'Color of shading on the face of command buttons
    vbButtonShadow = &H80000010     'Color of shading on the edge of command buttons
    vbGrayText = &H80000011    ' Grayed (disabled) text
    vbButtonText = &H80000012      ' Text color on push buttons
    vbInactiveCaptionText = &H80000013  'Color of text in an inactive caption
    vb3DHighlight = &H80000014      'Highlight color for 3D display elements
    vb3DDKShadow = &H80000015       'Darkest shadow color for 3D display elements
    vb3DLight = &H80000016     ' Second lightest of the 3D colors after vb3Dhighlight
    vb3DFace = &H8000000F       'Color of text face
    vb3DShadow = &H80000010     'Color of text shadow
    vbInfoText = &H80000017     'Color of text in ToolTips
    vbInfoBackground = &H80000018   'Background color of ToolTips
    vbViolet = &HFF8080
    vbVioletBright = &HFFC0C0
    vbForestGreen = &H228B22
    vbGray = &HE0E0E0
    vbLightBlue = &HFFD3A4
    vbLightGreen = &HABFCBD
    vbGreenLemon = &HB3FFBE
    vbYellowBright = &HC0FFFF
    vbOrange = &H2CCDFC
    vbBlack = 0
    vbBlue = &HFF0000
    vbCyan = &HFFFF00
    vbGreen = &HFF00
    vbMagenta = &HFF00FF
    vbRed = &HFF
    vbWhite = &HFFFFFF
    vbYellow = &HFFFF
End Enum
#If False Then
Private vbScrollBars, vbDesktop, vbActiveTitleBar, vbInactiveTitleBar, vbMenuBar, vbWindowBackground, vbWindowFrame, vbMenuText
Private vbWindowText, vbTitleBarText, vbActiveBorder, vbInactiveBorder, vbApplicationWorkspace, vbHighlight, vbHighlightText
Private vbButtonFace, vbButtonShadow, vbGrayText, vbButtonText, vbInactiveCaptionText, vb3DHighlight, vb3DDKShadow, vb3DLight
Private vb3DFace, vb3DShadow, vbInfoText, vbInfoBackground, vbViolet, vbVioletBright, vbForestGreen, vbGray, vbLightBlue, vbLightGreen
Private vbGreenLemon, vbYellowBright, vbOrange, vbBlack, vbBlue, vbCyan, vbGreen, vbMagenta, vbRed, vbWhite, vbYellow
#End If

Private Const BDR_RAISEDOUTER    As Long = &H1
Private Const BDR_SUNKENOUTER    As Long = &H2
Private Const BDR_RAISEDINNER    As Long = &H4
Private Const BDR_SUNKENINNER    As Long = &H8
Private Const EDGE_RAISED        As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_ETCHED        As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP          As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT            As Long = &H1
Private Const BF_TOP             As Long = &H2
Private Const BF_RIGHT           As Long = &H4
Private Const BF_BOTTOM          As Long = &H8
Private Const BF_FLAT            As Long = &H4000
Private Const BF_RECT            As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                                                qrc As TYPERECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Boolean

Public Sub PaintControl(picBox As PictureBox, _
                        Tampilan As Appearance, _
                        Optional ByVal prov_BackColor As Long, _
                        Optional ByVal prov_ForeColor As Long, _
                        Optional ByVal sCaption As String, _
                        Optional ByVal PDown As Boolean)

Dim typRect As TYPERECT
Dim origScaleMode

    On Error Resume Next

    With picBox
        .BorderStyle = 0
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .Cls
        .BackColor = prov_BackColor
        .ForeColor = prov_ForeColor
    End With 'picBox
    With typRect
        .Right = picBox.ScaleWidth
        .Top = picBox.ScaleTop
        .Left = picBox.ScaleLeft     '    .Top = picBox.ScaleWidth
        .Bottom = picBox.ScaleHeight
    End With
    Select Case Tampilan 'm_Appearance
    Case 0
        DrawEdge picBox.hdc, typRect, EDGE_BUMP, BF_FLAT ' BF_FLAT
    Case 1 'HalfRaised
        DrawEdge picBox.hdc, typRect, BDR_RAISEDINNER, BF_RECT 'HalfRaised
    Case 2 'Raised

        With picBox
            DrawEdge .hdc, typRect, EDGE_RAISED, BF_RECT
            picBox.Line (1, picBox.ScaleHeight - 2)-(picBox.ScaleWidth - 2, picBox.ScaleHeight - 2), vbApplicationWorkspace
            picBox.Line (0, picBox.ScaleHeight - 1)-(picBox.ScaleWidth, picBox.ScaleHeight - 1), vb3DDKShadow
 ' vbApplicationWorkspace
        End With 'picBox
    Case 3 'sunken
        DrawEdge picBox.hdc, typRect, BDR_SUNKENOUTER, BF_RECT
    Case 4 'etched
        DrawEdge picBox.hdc, typRect, EDGE_ETCHED, BF_RECT
    Case 5 'Bump
        DrawEdge picBox.hdc, typRect, EDGE_BUMP, BF_RECT
    End Select
    picBox.ScaleMode = origScaleMode
    If PDown Then
        picBox.CurrentX = ((picBox.ScaleWidth - picBox.TextWidth(sCaption)) / 2) + 1
        picBox.CurrentY = ((picBox.ScaleHeight - picBox.TextHeight(sCaption)) / 2) + 1
    Else
        picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(sCaption)) / 2
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(sCaption)) / 2
    End If
    picBox.Print sCaption
    If picBox.AutoRedraw Then
        picBox.Refresh
    End If
    On Error GoTo 0

End Sub
Public Sub Sleep(ByVal Seconds As Double)
Dim TempTime As Double
    TempTime = Timer
    Do While Timer - TempTime < Seconds
        DoEvents
        If Timer < TempTime Then
            TempTime = TempTime - 24# * 3600#
        End If
    Loop
End Sub

