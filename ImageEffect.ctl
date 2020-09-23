VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ImageEffect 
   BackColor       =   &H00000000&
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   1
      Top             =   0
      Width           =   6795
      Begin VB.PictureBox ImagePreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         FillColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3045
         Left            =   2040
         ScaleHeight     =   203
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   216
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.PictureBox picHidden1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3075
         Left            =   3120
         ScaleHeight     =   3045
         ScaleWidth      =   3240
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image imgBox1 
         Height          =   375
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgBox2 
         Height          =   375
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgtempt 
         Height          =   3045
         Left            =   3120
         Top             =   1320
         Visible         =   0   'False
         Width           =   3240
      End
   End
   Begin VB.PictureBox picTemplate 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   600
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1688
      Left            =   9600
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   1200
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   7680
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.PictureBox Picture300 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   7320
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   2610
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImageEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Cteated By Joko Mulyono
'Email:dantex_765@hotmail.com

Private HC                        As Long
Private m_BlockSize               As Integer
Private Const m_def_BlockSize     As Integer = 5
Private ms                        As Integer
Private Const STRETCH_HALFTONE    As Long = &H4&
Private Type POINTAPI
    X                                 As Long
    Y                                 As Long
End Type
Private Type TYPERECT
    Left                              As Long
    Top                               As Long
    Right                             As Long
    Bottom                            As Long
End Type
Private Type rBlendProps
    tBlendOp                          As Byte
    tBlendOptions                     As Byte
    tBlendAmount                      As Byte
    tAlphaType                        As Byte
End Type
Public Enum EffectLaser
    LaserLeft = 0
    LaserRight = 1
    LaserDown = 2
    LaserUp = 3
    LaserUp2 = 4
    LaserDown2 = 5
    LaserLB = 6
    LaserRB = 7
    LaserR13 = 8
    LaserCenter = 9
End Enum
#If False Then
Private LaserLeft, LaserRight, LaserDown, LaserUp, LaserUp2, LaserDown2, LaserLB, LaserRB, LaserR13
#End If

Private Const EXT                 As String = "JPG/BMP/WMF/GIF"
Private Const m_def_clearF        As Boolean = False
Private m_clearF                  As Boolean
Private X                         As Integer
Private Y                         As Integer
Private color1                    As Long
Private color2                    As Long
Private r                         As Integer
Private g                         As Integer
Private b                         As Integer
Private r2                        As Integer
Private g2                        As Integer
Private b2                        As Integer
Private Percent                   As Integer
Private m_FileName                As String
Private m_Value                   As Long
Private m_TransValue              As Integer
Private Const m_def_FileName      As String = ""
Private Const m_def_TransValue    As Integer = 1
Private Const m_def_Value         As Integer = 255
Private StartTime                 As Long
Private TotalDuration             As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, _
                                                 ByVal ySrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nXOrg As Long, _
                                                    ByVal nYOrg As Long, _
                                                    lpPt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, _
                                                        ByVal nStretchMode As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long




Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal hSrcDC As Long, _
                                                   ByVal xSrc As Long, _
                                                   ByVal ySrc As Long, _
                                                   ByVal widthSrc As Long, _
                                                   ByVal heightSrc As Long, _
                                                   ByVal blendFunct As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal length As Long)
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long
Private Declare Function timeGetTime Lib "winmm" () As Long
Private Sub BlindsHorizontal_Double()
Dim Stripes      As Integer
Dim I            As Integer
Dim j            As Integer
Dim StripeHeight As Integer
Dim ms           As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    StripeHeight = 20
    Stripes = Fix(picTemplate.ScaleHeight / StripeHeight)
    On Error Resume Next
    ms = TotalDuration / StripeHeight
    For j = 0 To Stripes
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, (Picture2.Width / 2 - picTemplate.Width / 2), I * StripeHeight, picTemplate.ScaleWidth, j, 0, I * StripeHeight, picTemplate.ScaleWidth, j, &HCC0020
            Picture2.PaintPicture picTemplate.Picture, (Picture2.Width / 2 - picTemplate.Width / 2), I * StripeHeight + 20, picTemplate.ScaleWidth, -j, 0, I * StripeHeight + 20, picTemplate.ScaleWidth, -j, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub BlindsHorizontal_Down()
Dim Stripes      As Integer
Dim I            As Integer
Dim j            As Integer
Dim StripeHeight As Integer
Dim ms           As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
' If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    StripeHeight = 20
    Stripes = Fix(picTemplate.ScaleHeight / StripeHeight)
    On Error Resume Next
    ms = TotalDuration / StripeHeight
    For j = 1 To StripeHeight
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, (Picture2.Width / 2 - picTemplate.Width / 2), I * StripeHeight, picTemplate.ScaleWidth, j, 0, I * StripeHeight, picTemplate.ScaleWidth, j, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub BlindsHorizontal_Up()
Dim Stripes      As Integer
Dim I            As Integer
Dim j            As Integer
Dim StripeHeight As Integer
Dim ms           As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    StripeHeight = 20
    Stripes = Fix(picTemplate.ScaleHeight / StripeHeight)
    On Error Resume Next
    ms = TotalDuration / StripeHeight
    For j = 1 To StripeHeight
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, (Picture2.Width / 2 - picTemplate.Width / 2), I * StripeHeight + 20, picTemplate.ScaleWidth, -j, 0, I * StripeHeight + 20, picTemplate.ScaleWidth, -j, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub BlindsVertical_Double()
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim I           As Integer
Dim j           As Integer
Dim ms          As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    Stripewidth = 20
    Stripes = (picTemplate.ScaleWidth) / Stripewidth
    On Error Resume Next
    ms = TotalDuration / Stripewidth
    For j = 1 To Stripewidth
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, I * Stripewidth, 0, j, picTemplate.ScaleHeight, I * Stripewidth, 0, j, picTemplate.ScaleHeight, &HCC0020
            Picture2.PaintPicture picTemplate.Picture, I * Stripewidth, 0, -j, picTemplate.ScaleHeight, I * Stripewidth, 0, -j, picTemplate.ScaleHeight, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub BlindsVertical_Left()
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim I           As Integer
Dim j           As Integer
Dim ms          As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    Stripewidth = 20
    Stripes = picTemplate.ScaleWidth / Stripewidth
    On Error Resume Next
    ms = TotalDuration / Stripewidth
    For j = 1 To Stripewidth
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, I * Stripewidth, 0, j, picTemplate.ScaleHeight, I * Stripewidth, 0, j, picTemplate.ScaleHeight, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub BlindsVertical_Right()
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim I           As Integer
Dim j           As Integer
Dim ms          As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    Stripewidth = 20
    Stripes = (picTemplate.ScaleWidth) / Stripewidth
    On Error Resume Next
    ms = TotalDuration / Stripewidth
    For j = 1 To Stripes
        Startplay
        For I = 0 To Stripes
            Picture2.PaintPicture picTemplate.Picture, I * Stripewidth + 20, 0, -j, picTemplate.ScaleHeight, I * Stripewidth + 20, 0, -j, picTemplate.ScaleHeight, &HCC0020
        Next I
        Endplay (ms)
    Next j
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Public Property Get BlockSize() As Integer
    BlockSize = m_BlockSize
End Property
Public Property Let BlockSize(ByVal New_BlockSize As Integer)
    m_BlockSize = New_BlockSize
    PropertyChanged "BlockSize"
End Property
Private Sub Centering()

'Created by someone ( I don't know the original resource)
Dim nY                     As Single
Dim arrRandom()            As Long
Dim arrRandomize()         As Long
Dim blksize                As Integer
Dim blksizeX               As Integer
Dim blksizeY               As Integer
Dim imgW                   As Variant
Dim imgH                   As Variant
Dim X                      As Long
Dim Y                      As Long
Dim Z                      As Long
Dim xWidthInBlocks         As Long
Dim xWidthInBlocksConstant As Long
Dim yWidthInBlocks         As Long
Dim zz                     As Long
Dim zcount                 As Long
    Set Picture2.Picture = Picture2.Image
    With picTemplate
        imgBox1.Picture = .Image 'LoadPicture("F:\Gambar\Hantu2.bmp") 'picture
        imgBox2.Picture = .Image 'LoadPicture("F:\Gambar\Hantu2.bmp") 'get 2nd picture
        .Picture = Picture2.Picture
        .PaintPicture imgBox1, 0, 0, , , , , , , vbSrcPaint
        Set .Picture = .Image
        .PaintPicture imgBox2, 0, 0, , , , , , , vbSrcAnd
        Set .Picture = .Image
    End With 'picTemplate
    Picture2.AutoRedraw = False
    imgW = imgBox1.Width
    imgH = imgBox1.Height  'set height and width variables
    If imgW + 20 > Picture2.ScaleWidth Then
        imgW = Picture2.ScaleWidth - 0
    End If
    If imgH + 20 > Picture2.ScaleHeight Then
        imgH = Picture2.ScaleHeight - 0
    End If
    blksize = 5
    ReDim arrRandom((imgW * imgH) / blksize, 4) As Long
    ReDim arrRandomize((imgW * imgH) / blksize) As Long
    Z = 1
    For Y = 0 To imgH Step blksize
        For X = 0 To imgW Step blksize
            blksizeX = blksize
            blksizeY = blksize
            If imgW - X = 0 Then
                GoTo lblNextx
            End If
            If imgW - X < blksize Then
                blksizeX = imgW - X
            End If
            If imgH - Y = 0 Then
                GoTo lblNexty
            End If
            If imgH - Y < blksize Then
                blksizeY = imgH - Y
            End If
            arrRandom(Z, 0) = Z
            arrRandom(Z, 1) = X
            arrRandom(Z, 2) = Y
            arrRandom(Z, 3) = blksizeX - 2
            arrRandom(Z, 4) = blksizeY - 2
            Z = Z + 1
lblNextx:
        Next X
lblNexty:
    Next Y
    Z = Z - 1
    For X = 1 To Z
        arrRandomize(X) = X
    Next X
DoNotRandomize:
    xWidthInBlocks = Int(imgW / blksize)
    If (imgW Mod blksize) > 0 Then
        xWidthInBlocks = xWidthInBlocks + 1
    End If
    xWidthInBlocksConstant = xWidthInBlocks
    yWidthInBlocks = Z / xWidthInBlocks
    yWidthInBlocks = yWidthInBlocks - 1
    zz = 0
    zcount = 0
lblCircleStart:
    For X = 1 To xWidthInBlocks
        zz = zz + 1
        zcount = zcount + 1
        GoSub lblLaser

    Next X
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For Y = 1 To yWidthInBlocks
        zz = zz + xWidthInBlocksConstant
        zcount = zcount + 1
        GoSub lblLaser

    Next Y
    yWidthInBlocks = yWidthInBlocks - 1
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For X = 1 To xWidthInBlocks - 1
        zz = zz - 1
        zcount = zcount + 1
        GoSub lblLaser

    Next X
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For Y = 1 To yWidthInBlocks
        zz = zz - xWidthInBlocksConstant
        zcount = zcount + 1
        GoSub lblLaser

    Next Y
    yWidthInBlocks = yWidthInBlocks - 1
    xWidthInBlocks = xWidthInBlocks - 2
    If zcount >= Z Then
        GoTo lblFinished
    End If
    GoTo lblCircleStart
lblLaser:
    DoEvents
    For nY = 1 To 300
        DoEvents
    Next nY
    GetPixel picTemplate.hdc, arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0

'r = color1 Mod 256
    Picture2.PaintPicture picTemplate, arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0, , , arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0, arrRandom(arrRandomize(zz), 3) + 2, arrRandom(arrRandomize(zz), 4) + 2
    Return

lblFinished:
    Picture2.AutoRedraw = False
    Set Picture2.Picture = picTemplate.Image
End Sub
Public Property Get ClearFirst() As Boolean
    ClearFirst = m_clearF
End Property
Public Property Let ClearFirst(ByVal new_CF As Boolean)
    m_clearF = new_CF
End Property
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Picture2.Cls
End Sub
Private Sub Endplay(N As Integer)
'bad code but usefull
    Do While timeGetTime() - StartTime < N

    Loop
End Sub
Private Sub EnlargeBottom()
Dim Y  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For Y = 1 To picTemplate.ScaleHeight Step I '3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next Y
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeCH()
Dim X  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        For X = 1 To .ScaleWidth / 2 Step I  '3
            Startplay
            .PaintPicture picTemplate.Picture, picTemplate.ScaleWidth / 2, 0, -X, picTemplate.ScaleHeight, picTemplate.ScaleWidth / 2, 0, -picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight, &HCC0020
            .PaintPicture picTemplate.Picture, picTemplate.ScaleWidth / 2, 0, X, picTemplate.ScaleHeight, picTemplate.ScaleWidth / 2, 0, picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight, &HCC0020
            Endplay (ms)
        Next X
    End With 'Picture2
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeClose()
Dim X  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth / 2 Step I '3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, X, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight, &HCC0020
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, picTemplate.ScaleWidth / 2, 0, picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeCloseH()
Dim Y  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For Y = 1 To picTemplate.ScaleHeight / 2 Step I
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, Y, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight / 2, &HCC0020
        Picture2.PaintPicture picTemplate.Picture, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, 0, picTemplate.ScaleHeight / 2, picTemplate.ScaleWidth, picTemplate.ScaleHeight / 2, &HCC0020
        Endplay (ms)
    Next Y
    Picture2.Width = picTemplate.Width
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeCV() 'OKE
Dim Y  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    With picTemplate
        Set .Picture = .Image
        ms = TotalDuration / .ScaleHeight
        For Y = 1 To .ScaleHeight / 2 Step I
            Startplay
            Picture2.PaintPicture picTemplate, 0, .ScaleHeight / 2, .ScaleWidth, -Y, 0, .ScaleHeight / 2, .ScaleWidth, -.ScaleHeight / 2, &HCC0020
            Picture2.PaintPicture .Picture, 0, .ScaleHeight / 2, .ScaleWidth, Y, 0, .ScaleHeight / 2, .ScaleWidth, .ScaleHeight / 2, &HCC0020
            Endplay (ms)
        Next Y
    End With 'picTemplate
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeLeft()
Dim X  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step I
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, X, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeRight()
Dim X  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step I
        Startplay
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub EnlargeTop()
Dim Y  As Integer
Dim ms As Integer
Dim I  As Integer
    On Error GoTo Pesan
    If UserControl.Width > 600 Then
        I = 5
    Else
        I = 3
    End If
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For Y = 1 To picTemplate.ScaleHeight Step I 'Step 3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, Y, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next Y
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub StripesVertical()
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim I           As Integer
Dim ms          As Integer
Dim p           As Integer

    On Error Resume Next
    Stripewidth = 40
    Stripes = Fix(picTemplate.ScaleWidth / Stripewidth)

    ms = (TotalDuration / Picture2.ScaleHeight) / 2
    For I = 1 To Picture2.ScaleHeight Step 2
        Startplay
        For p = 0 To Stripes
            Picture2.AutoRedraw = False
            Picture2.PaintPicture picTemplate.Picture, p * 80, -picTemplate.ScaleHeight + 1 + I, Stripewidth, picTemplate.ScaleHeight, 80 * p, 0, Stripewidth, picTemplate.ScaleHeight
            Picture2.PaintPicture picTemplate.Picture, (80 * p) + 40, picTemplate.ScaleHeight - I, Stripewidth, picTemplate.ScaleHeight, 80 * p + 40, 0, Stripewidth, picTemplate.ScaleHeight ' - I ', Picture1.ScaleHeight - I, 80 * P, Picture1.ScaleHeight, 40, I, &HCC0020
            If I >= Picture2.ScaleHeight + 5 Then
                Exit For
            End If
        Next p
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Picture
    On Error GoTo 0
End Sub
Private Sub StripesHorizontal()
Dim X  As Integer
Dim ms As Integer
Dim I As Integer
Dim StripeHeight  As Integer
Dim Stripes As Integer
On Error GoTo Pesan
ms = TotalDuration / picTemplate.ScaleWidth + 5
        StripeHeight = 40
        Stripes = Fix(picTemplate.ScaleHeight / StripeHeight)
        Picture2.AutoRedraw = False
        For X = 1 To Picture2.ScaleWidth Step 5
            Startplay
            For I = 0 To Stripes
                   Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - 5 - X, 80 * I, X + 5, StripeHeight, 0, 80 * I, X, StripeHeight 'code lain
                   Picture2.PaintPicture picTemplate.Picture, -picTemplate.ScaleWidth + 3 + X, (80 * I) + 40, picTemplate.ScaleWidth + 5, StripeHeight, 0, (80 * I) + 40, picTemplate.ScaleWidth, StripeHeight, &HCC0020
                   If X >= picTemplate.ScaleWidth Then
                           Exit For
                   End If
            Next I
            Endplay (ms)
        Next X
    Set Picture2.Picture = picTemplate.Picture
    'On Error GoTo 0
Pesan:
If Err.Number <> 0 Then
   If Err.Number = 5 Then
      Resume Next
   End If
End If
End Sub
Private Sub Chess1()
Dim PWidth      As Integer
Dim PHeight     As Integer
Dim I           As Integer
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim ms          As Integer
Dim pWidth2     As Integer
Dim p           As Integer
    On Error Resume Next
    Stripewidth = Fix(Picture2.ScaleWidth / 16)
    Stripes = Fix(Picture2.ScaleWidth / Stripewidth)
    For PWidth = 0 To Picture2.ScaleWidth Step (Stripewidth * 2)
        PHeight = 0
        ms = TotalDuration / Picture2.ScaleHeight
        For I = 1 To Picture2.ScaleHeight Step (Stripewidth * 2)
            Startplay
            For p = 0 To Stripes

                With Picture2
                    .AutoRedraw = False
                    .PaintPicture picTemplate.Picture, PWidth, I, Stripewidth, Stripewidth, PWidth, I, Stripewidth, Stripewidth ' picTemplate.ScaleHeight
                    .PaintPicture picTemplate.Picture, PWidth + Stripewidth, I + Stripewidth, Stripewidth, Stripewidth, PWidth + Stripewidth, I + Stripewidth, Stripewidth, Stripewidth  ' picTemplate.ScaleHeight
                End With 'Picture2
                PHeight = PHeight + 5
                If PHeight >= Picture2.ScaleHeight + 5 Then
                    Exit For
                End If
            Next p
            Endplay (ms)
        Next I
    Next PWidth
    For pWidth2 = 0 To Picture2.ScaleWidth Step (Stripewidth * 2)
        PHeight = 0
        ms = TotalDuration / Picture2.ScaleHeight
        For I = 1 To Picture2.ScaleHeight Step (Stripewidth * 2)
            Startplay
            For p = 0 To Stripes
                Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - pWidth2 - Stripewidth, I, Stripewidth, Stripewidth, picTemplate.ScaleWidth - pWidth2 - Stripewidth, I, Stripewidth, Stripewidth   ' picTemplate.ScaleHeight
                Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - pWidth2 - (Stripewidth * 2), I + Stripewidth, Stripewidth, Stripewidth, picTemplate.ScaleWidth - pWidth2 - (Stripewidth * 2), I + Stripewidth, Stripewidth, Stripewidth ' picTemplate.ScaleHeight
                PHeight = PHeight + 5
                If PHeight >= Picture2.ScaleHeight + 5 Then
                    Exit For
                End If
            Next p
            Endplay (ms)
        Next I
    Next pWidth2
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub Chess2()
Dim PWidth      As Integer
Dim PHeight     As Integer
Dim I           As Integer
Dim Stripewidth As Integer
Dim Stripes     As Integer
Dim ms          As Integer
Dim pWidth2     As Integer
Dim p           As Integer
    On Error Resume Next
    Stripewidth = Fix(Picture2.ScaleWidth / 16)
    Stripes = Fix(Picture2.ScaleWidth / Stripewidth)
    For PWidth = 0 To Picture2.ScaleWidth Step (Stripewidth * 2)
        PHeight = 0
        ms = TotalDuration / Picture2.ScaleHeight
        For I = 1 To Picture2.ScaleHeight Step (Stripewidth * 2)
            Startplay
            For p = 0 To Stripes

                With Picture2
                    .AutoRedraw = False
                    .PaintPicture picTemplate.Picture, PWidth, I, Stripewidth, Stripewidth, PWidth, I, Stripewidth, Stripewidth ' picTemplate.ScaleHeight
                    .PaintPicture picTemplate.Picture, PWidth + Stripewidth, I + Stripewidth, Stripewidth, Stripewidth, PWidth + Stripewidth, I + Stripewidth, Stripewidth, Stripewidth  ' picTemplate.ScaleHeight
                End With 'Picture2
                PHeight = PHeight + 5
                If PHeight >= Picture2.ScaleHeight + 5 Then
                    Exit For
                End If
            Next p
            Endplay (ms)
        Next I
    Next PWidth
    For pWidth2 = 0 To Picture2.ScaleWidth Step (Stripewidth * 2)
        PHeight = 0
        ms = TotalDuration / Picture2.ScaleHeight
        For I = 1 To Picture2.ScaleHeight Step (Stripewidth * 2)
            Startplay
            For p = 0 To Stripes
                Picture2.PaintPicture picTemplate.Picture, pWidth2 - Stripewidth, I, Stripewidth, Stripewidth, pWidth2 - Stripewidth, I, Stripewidth, Stripewidth  ' picTemplate.ScaleHeight
                Picture2.PaintPicture picTemplate.Picture, pWidth2 - (Stripewidth * 2), I + Stripewidth, Stripewidth, Stripewidth, pWidth2 - (Stripewidth * 2), I + Stripewidth, Stripewidth, Stripewidth ' picTemplate.ScaleHeight
                PHeight = PHeight + 5
                If PHeight >= Picture2.ScaleHeight + 5 Then
                    Exit For
                End If
            Next p
            Endplay (ms)
        Next I
    Next pWidth2
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Public Property Get FileName() As String
    FileName = m_FileName
    If LenB(m_FileName) = 0 Then Picture2.Picture = Nothing
End Property
Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    If LenB(m_FileName) = 0 Then
        Picture2.Picture = Nothing
    End If
    PropertyChanged "FileName"
End Property
Private Sub Fitting(picSource As PictureBox, _
                    picThumb As PictureBox)
Dim lLeft        As Long
Dim lTop         As Long
Dim lWidth       As Long
Dim lHeight      As Long
Dim lForeColor   As Long
Dim hBrush       As Long
Dim hDummyBrush  As Long
Dim lOrigMode    As Long
Dim fScale       As Single
Dim uBrushOrigPt As POINTAPI
    picThumb.BackColor = vbBlack 'ButtonFace
    picThumb.AutoRedraw = True
    picThumb.Cls
    If picSource.Width <= picThumb.Width - 2 And picSource.Height <= picThumb.Height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.Width > picSource.Height, (picThumb.Width - 2) / picSource.Width, (picThumb.Height - 2) / picSource.Height)
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    lLeft = Int((picThumb.Width - lWidth) / 2)
    lTop = Int((picThumb.Height - lHeight) / 2)
    lForeColor = picThumb.ForeColor
    lOrigMode = SetStretchBltMode(picThumb.hdc, STRETCH_HALFTONE)
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picThumb.hdc, hDummyBrush)
    UnrealizeObject hBrush
    With picThumb
        SetBrushOrgEx .hdc, lLeft, lTop, uBrushOrigPt
        hDummyBrush = SelectObject(.hdc, hBrush)
        StretchBlt .hdc, lLeft, lTop, lWidth, lHeight, picSource.hdc, 0, 0, picSource.Width, picSource.Height, vbSrcCopy
        SetStretchBltMode .hdc, lOrigMode
        hBrush = SelectObject(.hdc, hDummyBrush)
    End With 'picThumb
    UnrealizeObject hBrush
    SetBrushOrgEx picThumb.hdc, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt
    hDummyBrush = SelectObject(picThumb.hdc, hBrush)
    DeleteObject hDummyBrush
    picThumb.ForeColor = lForeColor
End Sub
Private Sub GetEffect(ByVal TheEffect As Integer)

    Picture2.Cls
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Select Case TheEffect
           Case 0
              NoEffect
           Case 1
              BlindsHorizontal_Double
           Case 2
              BlindsHorizontal_Down
           Case 3
              BlindsHorizontal_Up
           Case 4
              BlindsVertical_Double
           Case 5
              BlindsVertical_Left
           Case 6
              BlindsVertical_Right
           Case 7
              Centering
           Case 8
              Chess1
           Case 9
              Chess2
           Case 10
              EnlargeBottom
           Case 11
              EnlargeCH
           Case 12
              EnlargeClose
           Case 13
              EnlargeCloseH
           Case 14
              EnlargeCV
           Case 15
              EnlargeLeft
           Case 16
              EnlargeRight
           Case 17
              EnlargeTop
           Case 18
              ShowLaser (LaserCenter)
           Case 19
              ShowLaser (LaserDown) '
           Case 20
              ShowLaser (LaserDown2) '
           Case 21
              ShowLaser (LaserLB) '
           Case 22
              ShowLaser (LaserLeft) '
           Case 23
              ShowLaser (LaserR13) '
           Case 24
              ShowLaser (LaserRB) '
           Case 25
              ShowLaser (LaserRight) '
           Case 26
              ShowLaser (LaserUp) '
           Case 27
              ShowLaser (LaserUp2) '
           Case 28
              RandomBlock BlockSize
           Case 29
              RevealDown
           Case 30
              RevealUp
           Case 31
              SliceDown
           Case 32
              SliceHorizontal1
           Case 33
              SliceHorizontal2
           Case 34
              SliceHorizontal3
           Case 35
              SliceLeft
           Case 36
              SliceRight
           Case 37
              SliceUp
           Case 38
              SliceVertical1
           Case 39
              SliceVertical2
           Case 40
              SliceVertical3
           Case 41
              ShowSlide "down"
           Case 42
              ShowSlide "down2"
           Case 43
              ShowSlide "Left"
           Case 44
              ShowSlide "left2"
           Case 45
              ShowSlide "right"
           Case 46
              ShowSlide "right2"
           Case 47
              ShowSlide "up"
           Case 48
              ShowSlide "up2"
           Case 49
              Slide_BottomLeft
           Case 50
              Slide_BottomRight
           Case 51
              SlideLeft
           Case 52
              SlideRight
           Case 53
              SlideRight_RV
           Case 54
              SlideUp
           Case 55
              StripesHorizontal
           Case 56
              StripesVertical
           Case 57
              StretchClose
           Case 58
              StretchDown
           Case 59
              StretchDown_B
           Case 60
              StretchLeft
           Case 61
              StretchLeft_B
           Case 62
              StretchR
           Case 63
              StretchRight_B
           Case 64
              StretchUp
           Case 65
              StretchUp_B
           Case 66
              Translucent m_Value
           Case 67
              Transparent
           Case 68
              WaveDown
           Case 69
              WaveLeft
           Case 70
              WaveRight
           Case 71
              WaveUp
           Case 72
              WipesCenter
           Case 73
              WipesClose
           Case 74
              WipesCloseV
           Case 75
              WipesLeft
           Case 76
              WipesOpenH
           Case 77
              WipesRights
           Case 78
              Zoom_LeftBottom
           Case 79
              Zoom_RightBottom
           Case 80
              Zoom_UpLeft
           Case 81
              Zoom_UpRight
           Case 82
              ZoomIn
           Case 83
              ZoomLeft_B
           Case 84
              ZoomOut
    End Select
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
'Set Picture2.Picture = Picture2.Image
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
End Sub
Public Function GetEffectName(ByVal Fname As Integer) As String

Dim EName As String
    Select Case Fname
           Case 0
             EName = "NoEffect"
           Case 1
              EName = "BlindsHorizontal_Double"
           Case 2
              EName = "BlindsHorizontal_Down"
           Case 3
              EName = "BlindsHorizontal_Up"
           Case 4
              EName = "BlindsVertical_Double"
           Case 5
              EName = "BlindsVertical_Left"
           Case 6
              EName = "BlindsVertical_Right"
           Case 7
              EName = "Centering"
           Case 8
              EName = "Chess1"
           Case 9
              EName = "Chess2"
           Case 10
              EName = "EnlargeBottom"
           Case 11
              EName = "EnlargeCH"
           Case 12
              EName = "EnlargeClose"
           Case 13
              EName = "EnlargeCloseH"
           Case 14
              EName = "EnlargeCV"
           Case 15
              EName = "EnlargeLeft"
           Case 16
              EName = "EnlargeRight"
           Case 17
              EName = "EnlargeTop"
           Case 18
              EName = "LaserCenter"
           Case 19
              EName = "LaserDown"
           Case 20
              EName = "LaserDown2"
           Case 21
              EName = "LaserLB"
           Case 22
              EName = "LaserLeft"
           Case 23
              EName = "LaserR13"
           Case 24
              EName = "LaserRB"
           Case 25
              EName = "LaserRight"
           Case 26
              EName = "LaserUp"
           Case 27
              EName = "LaserUp2"
           Case 28
              EName = "RandomBlock"
           Case 29
              EName = "RevealDown"
           Case 30
              EName = "RevealUp"
           Case 31
              EName = "SliceDown"
           Case 32
              EName = "SliceHorizontal1"
           Case 33
              EName = "SliceHorizontal2"
           Case 34
              EName = "SliceHorizontal3"
           Case 35
              EName = "SliceLeft"
           Case 36
              EName = "SliceRight"
           Case 37
              EName = "SliceUp"
           Case 38
              EName = "SliceVertical1"
           Case 39
              EName = "SliceVertical2"
           Case 40
              EName = "SliceVertical3"
           Case 41
              EName = "Slide down"
           Case 42
              EName = "Slide down2"
           Case 43
              EName = "Slide Left"
           Case 44
              EName = "Slide left2"
           Case 45
              EName = "Slide Right"
           Case 46
              EName = "Slide right2"
           Case 47
              EName = "Slide up"
           Case 48
              EName = "Slide up2"
           Case 49
              EName = "Slide_BottomLeft"
           Case 50
              EName = "Slide_BottomRight"
           Case 51
              EName = "SlideLeft"
           Case 52
              EName = "SlideRight"
           Case 53
              EName = "SlideRight_RV"
           Case 54
              EName = "SlideUp"
           Case 55
              EName = "StripesHorizontal"
           Case 56
              EName = "StripesVertical"
           Case 57
              EName = "StretchClose"
           Case 58
              EName = "StretchDown"
           Case 59
              EName = "StretchDown_B"
           Case 60
              EName = "StretchLeft"
           Case 61
              EName = "StretchLeft_B"
           Case 62
              EName = "StretchR"
           Case 63
              EName = "StretchRight_B"
           Case 64
              EName = "StretchUp"
           Case 65
              EName = "StretchUp_B"
           Case 66
              EName = "Translucent"
           Case 67
              EName = "Transparent"
           Case 68
              EName = "WaveDown"
           Case 69
              EName = "WaveLeft"
           Case 70
              EName = "WaveRight"
           Case 71
              EName = "WaveUp"
           Case 72
              EName = "WipesCenter"
           Case 73
              EName = "WipesClose"
           Case 74
              EName = "WipesCloseV"
           Case 75
              EName = "WipesLeft"
           Case 76
              EName = "WipesOpenH"
           Case 77
              EName = "WipesRights"
           Case 78
              EName = "Zoom_LeftBottom"
           Case 79
              EName = "Zoom_RightBottom"
           Case 80
              EName = "Zoom_UpLeft"
           Case 81
              EName = "Zoom_UpRight"
           Case 82
              EName = "ZoomIn"
           Case 83
              EName = "ZoomLeft_B"
           Case 84
              EName = "ZoomOut"
           
            
    End Select
    GetEffectName = EName
End Function
Private Function GetExtension(ByVal FullFilePath As String) As String
Dim p As Long
    If Len(FullFilePath) > 0 Then
        p = InStrRev(FullFilePath, ".")
        If p > 0 Then
            If p < Len(FullFilePath) Then
                GetExtension = Mid$(FullFilePath, p + 1)
            End If
        End If
    End If
End Function
Private Sub getRGBCOLORpixel(obj As Object, _
                             X As Integer, _
                             Y As Integer)
    color1 = GetPixel(obj.hdc, X, Y)
    r = color1 Mod 256
    b = Int(color1 / 65536)
    g = (color1 - (b * 65536) - r) / 256
End Sub
Private Sub SliceHorizontal1() '#24 'Horizontal Allez-Retour une
'By Joko Mulyono \ DANTE '09\09\2003
Dim X       As Integer
Dim Y       As Integer
Dim ms      As Integer
Dim PHeight As Integer

    With Picture2
        .AutoRedraw = False
        ms = TotalDuration / .ScaleWidth + 5
        For Y = 1 To .ScaleHeight + 41 Step 80
            PHeight = 40
            For X = 1 To .ScaleWidth + 5 Step 5
                Startplay
'mode #1
                .PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, Y + 40, X, 40, .ScaleWidth - X, Y + 40, X, 40
                .PaintPicture picTemplate.Picture, 0, Y, X, 40, 0, Y, X, 40 'oke
                PHeight = PHeight + 40
                Endplay (ms)
            Next X
        Next Y
    End With 'Picture2
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub SliceHorizontal2() '#25 'Horizontal Allez-Retour deuxieme
'By Joko Mulyono \ DANTE '09\09\2003
Dim X            As Integer
Dim Y            As Integer
Dim ms           As Integer
Dim oldwidth     As Integer
Dim newwidth     As Integer
Dim StripeHeight As Integer
    On Error GoTo Pesan
    StripeHeight = Int(Fix(Picture2.ScaleHeight / 12))
    oldwidth = picTemplate.ScaleWidth
    If oldwidth < Picture2.ScaleWidth Then
        newwidth = picTemplate.ScaleWidth
    Else
        newwidth = Picture2.ScaleWidth
    End If
'Bug: If new picture smaller then before, then can not make image well
'Fix: 20/07/2004
    Picture2.AutoRedraw = False
    ms = TotalDuration / newwidth
    For Y = 0 To Picture2.ScaleHeight + StripeHeight Step (StripeHeight * 2)
        For X = 1 To newwidth + 2 Step 5
            Startplay
            Picture2.PaintPicture picTemplate.Picture, Picture2.ScaleWidth - X + 1, Y, X, StripeHeight, 0, Y, X, StripeHeight
            Picture2.PaintPicture picTemplate.Picture, -Picture2.ScaleWidth + X - 1, Y + StripeHeight, Picture2.ScaleWidth, StripeHeight, 0, Y + StripeHeight, Picture2.ScaleWidth, StripeHeight, &HCC0020
            Endplay (ms)
        Next X
    Next Y
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
Pesan:
    If Err.Number <> 0 Then
        If Err.Number = 5 Then
            Resume Next
        End If
    End If
End Sub
Private Sub SliceHorizontal3()
'By Joko Mulyono \ DANTE '09\09\2003
Dim X            As Integer
Dim Y            As Integer
Dim ms           As Integer
Dim PKL          As Integer
Dim PHeight      As Integer
Dim StripeHeight As Integer
StripeHeight = Fix(Picture2.ScaleHeight / 12)

    On Error GoTo Pesan
    Picture2.AutoRedraw = False
    ms = TotalDuration / picTemplate.ScaleWidth
    For Y = 0 To Picture2.ScaleHeight Step StripeHeight * 2
        For X = 1 To picTemplate.ScaleWidth + 5 Step 5
            Startplay
            Picture2.PaintPicture picTemplate.Picture, 1 + (picTemplate.ScaleWidth) - X, Y, X, StripeHeight, 0, Y, X, StripeHeight   'code lain
            If X >= picTemplate.ScaleWidth Then
                PHeight = PHeight + 1
                Exit For
            End If
            Endplay (ms)
        Next X
        For PKL = 1 To picTemplate.ScaleWidth + 5 Step 5
            Startplay
            Picture2.PaintPicture picTemplate.Picture, -1 + (-picTemplate.ScaleWidth) + PKL, Y + StripeHeight, picTemplate.ScaleWidth, StripeHeight, 0, Y + StripeHeight, picTemplate.ScaleWidth, StripeHeight, &HCC0020
            If PKL >= picTemplate.ScaleWidth Then
                Exit For
            End If
            Endplay (ms)
        Next PKL
        If PHeight > 6 Then
            Set Picture2.Picture = picTemplate.Image
            Exit Sub

        End If
    Next Y
Pesan:
    If Err.Number <> 0 Then
        If Err.Number = 5 Then
            Resume Next
        End If
    End If
End Sub
Private Sub ImagePreview_Click()
    Showme
End Sub
Private Sub Label1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub Laser_Center()

Dim nY                     As Integer
Dim arrRandom()            As Long
Dim arrRandomize()         As Long
Dim blksize                As Integer
Dim blksizeX               As Integer
Dim blksizeY               As Integer
Dim imgW                   As Variant
Dim imgH                   As Variant
Dim X                      As Long
Dim Y                      As Long
Dim Z                      As Long
Dim color1                 As Long
Dim r                      As Long
Dim g                      As Long
Dim b                      As Long
Dim xWidthInBlocks         As Long
Dim xWidthInBlocksConstant As Long
Dim yWidthInBlocks         As Long
Dim zz                     As Long
Dim zcount                 As Long
    Set Picture2.Picture = Picture2.Image
    With picTemplate
        imgBox1.Picture = .Image
        imgBox2.Picture = .Image
        .Picture = Picture2.Picture
        .PaintPicture imgBox1, 0, 0, , , , , , , vbSrcPaint
        Set .Picture = .Image
        .PaintPicture imgBox2, 0, 0, , , , , , , vbSrcAnd
        Set .Picture = .Image
    End With 'picTemplate
    Picture2.AutoRedraw = False
    imgW = imgBox1.Width
    imgH = imgBox1.Height  'set height and width variables
'If image bigger than pixMain
'or goes past pixMain's right border then clip accordingly
    If imgW + 20 > Picture2.ScaleWidth Then
        imgW = Picture2.ScaleWidth - 0
    End If
    If imgH + 20 > Picture2.ScaleHeight Then
        imgH = Picture2.ScaleHeight - 0
    End If
    blksize = 5 'size of transition squares
    ReDim arrRandom((imgW * imgH) / blksize, 4) As Long
'width * height divided by minimum blksize (maximum array size)576/10 * 432/10
    ReDim arrRandomize((imgW * imgH) / blksize) As Long
    Z = 1
    For Y = 0 To imgH Step blksize
        For X = 0 To imgW Step blksize
            blksizeX = blksize
            blksizeY = blksize
'detect last X block and skip if 0
            If imgW - X = 0 Then
                GoTo lblNextx
            End If
'detect last X block if smaller
            If imgW - X < blksize Then
                blksizeX = imgW - X
            End If
'detect last Y block and skip if 0
            If imgH - Y = 0 Then
                GoTo lblNexty
            End If
'detect last Y block if smaller
            If imgH - Y < blksize Then
                blksizeY = imgH - Y
            End If
            arrRandom(Z, 0) = Z
            arrRandom(Z, 1) = X
            arrRandom(Z, 2) = Y
            arrRandom(Z, 3) = blksizeX - 2 'originally -2 was not included here which seemed
            arrRandom(Z, 4) = blksizeY - 2 'to makes blocks a couple pixels bigger
            Z = Z + 1
lblNextx:
        Next X
lblNexty:
    Next Y
    Z = Z - 1
    For X = 1 To Z
        arrRandomize(X) = X
    Next X
DoNotRandomize:
    xWidthInBlocks = Int(imgW / blksize)
    If (imgW Mod blksize) > 0 Then
        xWidthInBlocks = xWidthInBlocks + 1
    End If
    xWidthInBlocksConstant = xWidthInBlocks
    yWidthInBlocks = Z / xWidthInBlocks
    yWidthInBlocks = yWidthInBlocks - 1
    zz = 0
    zcount = 0
lblCircleStart:
    For X = 1 To xWidthInBlocks
        zz = zz + 1
        zcount = zcount + 1
        GoSub lblLaser

    Next X
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For Y = 1 To yWidthInBlocks
        zz = zz + xWidthInBlocksConstant
        zcount = zcount + 1
        GoSub lblLaser

    Next Y
    yWidthInBlocks = yWidthInBlocks - 1
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For X = 1 To xWidthInBlocks - 1
        zz = zz - 1
        zcount = zcount + 1
        GoSub lblLaser

    Next X
    If zcount >= Z Then
        GoTo lblFinished
    End If
    For Y = 1 To yWidthInBlocks
        zz = zz - xWidthInBlocksConstant
        zcount = zcount + 1
        GoSub lblLaser

    Next Y
    yWidthInBlocks = yWidthInBlocks - 1
    xWidthInBlocks = xWidthInBlocks - 2
    If zcount >= Z Then
        GoTo lblFinished
    End If
    GoTo lblCircleStart
lblLaser:
    DoEvents
    For nY = 1 To 300
        DoEvents
    Next nY
    color1 = GetPixel(picTemplate.hdc, arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0)
    r = color1 Mod 256
    b = Int(color1 / 65536)
    g = (color1 - (b * 65536) - r) / 256
    Picture2.Line (picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight / 2)-(arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0), RGB(r, g, b)
    Picture2.PaintPicture picTemplate, arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0, , , arrRandom(arrRandomize(zz), 1) + 0, arrRandom(arrRandomize(zz), 2) + 0, arrRandom(arrRandomize(zz), 3) + 2, arrRandom(arrRandomize(zz), 4) + 2
    Return

lblFinished:
    Picture2.AutoRedraw = False
    Set Picture2.Picture = picTemplate.Image
End Sub
Public Sub LaserII()
    NoEffect

End Sub
Public Sub LoadGambar(ByVal FileName As String, _
                      Effect As Integer)
Dim sx As Single
Dim sy As Single
'    On Error GoTo Pesan
    Picture2.Cls
    picTemplate.Cls
    Picture2.Cls
    Label1.Caption = FileName & " \ " & GetEffectName(Effect)
    If Effect > 24 And Effect < 32 Then
        Picture2.Refresh
    Else
        If Not m_clearF Then
            Picture2.Picture = Nothing
        Else
            Picture2.AutoRedraw = True
            Set Picture2.Picture = Picture2.Image
        End If
'picTemplate.Picture = Nothing
'Picture2.Picture = Nothing
    End If
    Resizeme
    ImagePreview.AutoRedraw = True
    Picture3.AutoSize = True
    Picture3.Picture = LoadPicture(FileName)
'need to shrink
' If Picture3.Width > UserControl.Width / 15 Or Picture3.Height > UserControl.Height / 15 Then
    Fitting Picture3, picTemplate
    If Picture3.Width > Picture3.Height Then
        Shrink (Picture2.Width)
    ElseIf Picture3.Width < Picture3.Height Then
        Shrink (Picture2.Height)
    End If
    With picHidden1
        .Width = imgtempt.Width
        .Height = imgtempt.Height
        .Picture = imgtempt.Picture
    End With 'picHidden1
    With picTemplate
        .Cls
        .Picture = LoadPicture()
        sx = .ScaleWidth / 2 - ImagePreview.Width / 2
        sy = .ScaleHeight / 2 - ImagePreview.Height / 2 ' - Picture1.Height / 2
        .PaintPicture ImagePreview, sx, sy, ImagePreview.Width, ImagePreview.Height, 0, 0, ImagePreview.ScaleWidth, ImagePreview.ScaleHeight
        Set .Picture = .Image
    End With 'picTemplate
'Else '
'not need to srink
'Fitting Picture3, picTemplate
'End If
    GetEffect Effect
    If LenB(m_FileName) = 0 Then
        Picture2.Picture = Nothing
    End If
'Pesan:
'    If Err.Number <> 0 Then
'        If Err.Number = 481 Then
'            Picture2.Cls
'            Picture2.Picture = Nothing
'            Exit Sub
'
'        End If
'        'MsgBox Err.Number & " ," & Err.Description, vbCritical + vbOKOnly
'        'Exit Sub
'
'    End If
End Sub
Private Sub MakeTransparent(picBox1 As PictureBox, _
                            PicBox2 As PictureBox, _
                            picDest As PictureBox, _
                            ByVal Percent As Integer, _
                            ByVal XTimes As Integer)
Dim XT As Integer
    picDest.Width = picBox1.Width
    picDest.Height = picBox1.Height
    For XT = 0 To XTimes - 1
        For X = 0 To picBox1.ScaleWidth - 1
            For Y = 0 To picBox1.ScaleHeight - 1
                color1 = GetPixel(picBox1.hdc, X, Y)
                r = color1 Mod 256
                b = Int(color1 / 65536)
                g = (color1 - (b * 65536) - r) / 256
                color2 = GetPixel(PicBox2.hdc, X, Y)
                r2 = color2 Mod 256
                b2 = Int(color2 / 65536)
                g2 = (color2 - (b2 * 65536) - r2) / 256
                r = (((100 - Percent) * r) + (Percent * r2)) / 100
                g = (((100 - Percent) * g) + (Percent * g2)) / 100
                b = (((100 - Percent) * b) + (Percent * b2)) / 100
                SetPixel picDest.hdc, X, Y, RGB(r, g, b)
            Next Y
            If X Mod 10 = 0 Then
                picDest.Refresh
            End If
        Next X
    Next XT
    picDest.Refresh
End Sub
Private Sub NoEffect()
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub PESANUMUM()
    If Err.Number <> 0 Then
        If Err.Number = 5 Then
            On Error GoTo 0
        End If
    End If
End Sub
Private Sub Picture2_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    HC = ShowCursor(True)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    End
End Sub
Private Sub Picture2_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
'HC = ShowCursor(False)
End Sub
Private Sub Picture2_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub RandomBlock(SzBlock As Integer)

Dim nY             As Single
Dim arrRandom()    As Long
Dim arrRandomize() As Long
Dim blksize        As Integer
Dim blksizeX       As Integer
Dim blksizeY       As Integer
Dim imgW           As Variant
Dim imgH           As Variant
Dim X              As Long
Dim Y              As Long
Dim Z              As Long
Dim r1             As Long
Dim rr1            As Long
Dim r2             As Long
Dim rr2            As Long
Dim flgRandomize   As Boolean
    SzBlock = m_BlockSize
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    flgRandomize = True
    With picTemplate
        imgBox1.Picture = .Image 'LoadPicture("F:\Gambar\Hantu2.bmp") 'picture
        imgBox2.Picture = .Image 'LoadPicture("F:\Gambar\Hantu2.bmp") 'get 2nd picture
        .PaintPicture imgBox1, 0, 0, , , , , , , vbSrcPaint
        Set .Picture = .Image
        .PaintPicture imgBox2, 0, 0, , , , , , , vbSrcAnd
        Set .Picture = .Image
    End With 'picTemplate
    Picture2.AutoRedraw = False
    imgW = imgBox1.Width
    imgH = imgBox1.Height  'set height and width variables
    If imgW + 0 > Picture2.ScaleWidth Then
        imgW = Picture2.ScaleWidth - 0
    End If
    If imgH + 0 > Picture2.ScaleHeight Then
        imgH = Picture2.ScaleHeight - 0
    End If
    blksize = SzBlock
    ReDim arrRandom((imgW * imgH) / blksize, 4) As Long
    ReDim arrRandomize((imgW * imgH) / blksize) As Long
    Z = 1
    For Y = 0 To imgH Step blksize
        For X = 0 To imgW Step blksize
            blksizeX = blksize
            blksizeY = blksize
'detect last X block and skip if 0
            If imgW - X = 0 Then
                GoTo lblNextx
            End If
'detect last X block if smaller
            If imgW - X < blksize Then
                blksizeX = imgW - X
            End If
'detect last Y block and skip if 0
            If imgH - Y = 0 Then
                GoTo lblNexty
            End If
'detect last Y block if smaller
            If imgH - Y < blksize Then
                blksizeY = imgH - Y
            End If
            arrRandom(Z, 0) = Z
            arrRandom(Z, 1) = X
            arrRandom(Z, 2) = Y
            arrRandom(Z, 3) = blksizeX - 2 'originally -2 was not included here which seemed
            arrRandom(Z, 4) = blksizeY - 2 'to makes blocks a couple pixels bigger
            Z = Z + 1
lblNextx:
        Next X
lblNexty:
    Next Y
    Z = Z - 1
'Randomize
    For X = 1 To Z
        arrRandomize(X) = X  ' 1st load array sequentially
    Next X
    If Not flgRandomize Then
        GoTo DoNotRandomize
    End If
    For Y = 1 To 2 'Make two passes of randomize
        For X = 1 To Z
            r1 = Int((Rnd * Z)) + 1
            r2 = Int((Rnd * Z)) + 1 'generate two random #'s between 1 and z
            rr1 = arrRandomize(r1)
            rr2 = arrRandomize(r2)
            arrRandomize(r1) = rr2
            arrRandomize(r2) = rr1  'swap values between 2 rnd array slots
        Next X
    Next Y
DoNotRandomize:
    For X = 1 To Z
        For nY = 1 To 300
            DoEvents
        Next nY
        Picture2.PaintPicture picTemplate, arrRandom(arrRandomize(X), 1) + 0, arrRandom(arrRandomize(X), 2) + 0, , , arrRandom(arrRandomize(X), 1) + 0, arrRandom(arrRandomize(X), 2) + 0, arrRandom(arrRandomize(X), 3) + 2, arrRandom(arrRandomize(X), 4) + 2
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub Resizeme()
Dim Result As Single

    Result = UserControl.ScaleWidth / 4
    UserControl.Height = (UserControl.ScaleWidth - Result) * 15
    ImagePreview.Width = UserControl.Width / 15
    ImagePreview.Height = UserControl.Height / 15
    Picture2.Width = UserControl.Width / 15
    Picture2.Height = UserControl.Height / 15
    Picture3.Width = UserControl.Width / 15
    Picture3.Height = UserControl.Height / 15
    picTemplate.Width = UserControl.Width / 15
    picTemplate.Height = UserControl.Height / 15
    picEffect.Width = UserControl.Width / 15
    picEffect.Height = UserControl.Height / 15
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
End Sub
Private Sub RevealDown()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    On Error GoTo Pesan
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    PWidth = picTemplate.ScaleWidth
    PHeight = 1
    ms = (TotalDuration / picTemplate.ScaleHeight)
    For I = 1 To picTemplate.ScaleHeight / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, PWidth, PHeight, 0, 0, PWidth, PHeight, &HCC0020
        PHeight = PHeight + 2
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub RevealUp()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    On Error GoTo Pesan
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    PWidth = picTemplate.Width
    PHeight = 1
    ms = TotalDuration / picTemplate.ScaleHeight
    For I = 1 To picTemplate.ScaleHeight / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, (picTemplate.ScaleHeight - PHeight), PWidth, PHeight, 0, (picTemplate.ScaleHeight - PHeight), PWidth, PHeight, &HCC0020
        PHeight = PHeight + 2
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub

Public Sub ShowLaser(ByVal Orientation As EffectLaser)  'String)

Dim X As Integer
Dim Y As Integer

    On Error Resume Next
    Picture2.Cls
    Picture2.AutoRedraw = True
    Picture2.Width = picTemplate.Width
    Picture2.Height = picTemplate.Height
'If m_FileName = "" Then Picture2.Picture = Nothing
    Select Case Orientation
' Orientation=
'================
    Case LaserLeft
'================
        With picTemplate
            For X = .Width + 1 To 0 Step -1
                For Y = 0 To .Height
                    getRGBCOLORpixel picTemplate, X, Y
                    Picture2.Line (-1, (Picture2.ScaleHeight / 2))-(X - 1, Y), RGB(r, g, b)
                    SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
                Next Y
                Picture2.Refresh
            Next X
        End With 'picTemplate
'================
    Case LaserLB '"Laser LeftBottom"
'================
        With picTemplate
            For X = .Width + 1 To 0 Step -1
                For Y = 0 To .Height
                    getRGBCOLORpixel picTemplate, X, Y
                    Picture2.Line (-1, (Picture2.ScaleHeight))-(X - 1, Y), RGB(r, g, b)
                    SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
                Next Y
                Picture2.Refresh
            Next X
        End With 'picTemplate
'--------------------
'    Case "Left-Right"
'---------------------
'        With picTemplate
'            For X = .Width + 1 To 0 Step -1
'                For Y = 0 To .Height
'                    getRGBCOLORpixel picTemplate, X / 2, Y
'                    Picture2.Line (-1, (Picture2.ScaleHeight / 2))-(X / 2 - 1, Y), RGB(r, g, b)
'                    SetPixel Picture2.hdc, X / 2, Y, RGB(r, g, b) 'RGB(b, g, r)
'
'                Next Y
'                Picture2.Refresh
'
'            Next X
'        End With 'picTemplate
'=================
    Case LaserRight
'=================
        With picTemplate
            For X = 0 To .Width
                For Y = 0 To .Height
                    getRGBCOLORpixel picTemplate, X, Y
                    Picture2.Line ((Picture2.ScaleWidth), (Picture2.ScaleHeight / 2))-(X + 1, Y - 1), RGB(r, g, b)
                    SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
                Next Y
                Picture2.Refresh
            Next X
        End With 'picTemplate
'=================
    Case LaserRB '"Laser RightBottom"
'=================
        With picTemplate
            For X = 0 To .Width
                For Y = 0 To .Height
                    getRGBCOLORpixel picTemplate, X, Y
                    Picture2.Line ((Picture2.ScaleWidth), (Picture2.ScaleHeight))-(X + 1, Y - 1), RGB(r, g, b)
                    SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
                Next Y
                Picture2.Refresh
            Next X
        End With 'picTemplate
'===============
    Case LaserUp
'===============
        With picTemplate
            For Y = .Height - 1 To 0 Step -1
                For X = 0 To .Width - 1
                    getRGBCOLORpixel picTemplate, X, Y
                    Picture2.Line ((Picture2.ScaleWidth / 2), -1)-(X - 2, Y + 2), RGB(r, g, b)
                    SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
                Next X
                Picture2.Refresh
            Next Y
        End With 'picTemplate
    Case LaserDown
        For Y = 0 To picTemplate.Height
            For X = 0 To picTemplate.Width
                getRGBCOLORpixel picTemplate, X, Y
                Picture2.Line ((Picture2.ScaleWidth / 2), Picture2.ScaleHeight)-(X - 2, Y + 2), RGB(r, g, b)
                SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
            Next X
            Picture2.Refresh
        Next Y
'=====================================================
    Case LaserDown2 'laserDown [full] '01/05/2003 'by DANTE
'=====================================================
        For Y = 0 To picTemplate.Height
            For X = 0 To picTemplate.Width
                getRGBCOLORpixel picTemplate, X, Y
                Picture2.Line ((X), (Picture2.ScaleHeight))-(X + 1, Y), RGB(r, g, b)
' Picture2.Line ((-X), 0)-(-X + 1, Y + 2), RGB(r, g, b)
                SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
            Next X
            Picture2.Refresh
        Next Y
'====================
    Case LaserR13 '"R13"
'====================
        For Y = 0 To picTemplate.Height
            For X = 0 To picTemplate.Width
                getRGBCOLORpixel picTemplate, X, Y
                Picture2.Line ((-X), (Picture2.ScaleHeight))-(X, Y), RGB(r, g, b)
'SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
            Next X
            Picture2.Refresh
        Next Y
'=====================================================
    Case LaserUp2 'laserUp [full] '01/05/2003 'by DANTE
'=====================================================
        For Y = picTemplate.Height To 0 Step -1
            For X = 0 To picTemplate.Width
                getRGBCOLORpixel picTemplate, X, Y
                Picture2.Line ((X), ((-Picture2.ScaleHeight + Y)))-(X + 1, Y), RGB(r, g, b)
                SetPixel Picture2.hdc, X, Y, RGB(r, g, b)
            Next X
            Picture2.Refresh
        Next Y
     Case LaserCenter
          Laser_Center
    End Select
    Picture2.AutoRedraw = False
    Picture2.Picture = Picture2.Image
    On Error GoTo 0
End Sub
Private Sub Showme()
    Picture3.Picture = LoadResPicture(101, 0) ', 0  'App.Path & "\Logo2.bmp", 0
    Fitting Picture3, Picture2
End Sub
Public Sub ShowSave()

Dim I As String
Dim Z As String

    Z = App.Path
    I = InputBox(m_FileName, "Save picture", Z & "\ImageEffect.bmp")
    SavePicture Picture2.Image, I
End Sub
Private Sub ShowSlide(ByVal CommandStr As String)

Dim nX   As Integer
Dim nY   As Integer
Dim imgW As Integer
Dim imgH As Integer
Dim sx   As Single
Dim sy   As Single
    sx = picTemplate.ScaleWidth / 2 - ImagePreview.Width / 2
    sy = picTemplate.ScaleHeight / 2 - ImagePreview.Height / 2
    With UserControl
        .Picture2.Picture = .Picture2.Image
        .Picture2.AutoRedraw = False
        .imgtempt.Picture = .picTemplate.Picture
        If .imgtempt.Width > .Picture2.Width Then
            .imgtempt.Width = .Picture2.Width
        End If
        If .imgtempt.Height > .Picture2.Height Then
            .imgtempt.Height = .Picture2.Height
        End If
        imgW = .picTemplate.Width
        imgH = .picTemplate.Height
        .picHidden1.Picture = .Picture2.Picture
        .picHidden1.PaintPicture .picTemplate.Image, 0, 0
        .picHidden1.Picture = .picTemplate.Image
' 1 = Slide   2 = Push
        Select Case CommandStr
'=====================
        Case "right", "right1"
'=====================
            .picTemplate.Width = imgW
            .picTemplate.Height = imgH
'.picTemplate.PaintPicture .imgtempt, 0, 0
            .picTemplate.Picture = .picTemplate.Image
            For nX = imgW - 1 To 1 Step -1
                .Picture2.PaintPicture .picTemplate.Image, 0, 0, , , nX, 0, imgW, imgH
                For nY = 1 To 2000
                    DoEvents
                Next nY
            Next nX
'==============
        Case "right2"
'==============
            .picTemplate.Width = imgW * 2
            .picTemplate.Height = imgH
            .picTemplate.PaintPicture .Picture2.Picture, imgW, 0, , , 0, 0, imgW, imgH
'.picTemplate.PaintPicture .imgtempt, 0, 0
            .picTemplate.Picture = .picTemplate.Image
            For nX = imgW To 1 Step -1
                .Picture2.PaintPicture .picTemplate.Picture, 0, 0, , , nX, 0, imgW, imgH
                For nY = 1 To 2000
                    DoEvents
                Next nY
            Next nX
            .picTemplate.Width = imgW  'normalize width
'=====================
        Case "left", "left1"
'=====================
            .picTemplate.Width = imgW
            .picTemplate.Height = imgH
            Set picTemplate.Picture = picTemplate.Image
            .picTemplate.PaintPicture .picTemplate, 0, 0
            .picTemplate.Picture = .picTemplate.Image
            For nX = 1 To imgW Step 1
                .Picture2.PaintPicture .picTemplate.Picture, 0 + (imgW - nX), 0, , , 0, 0, nX, imgH
                For nY = 1 To 2000
                    DoEvents
                Next nY
            Next nX
'=======================================================
        Case "left2" 'fix:20/09/2004
'=======================================================
            .picTemplate.Width = imgW * 2
            .picTemplate.Height = imgH
            .picTemplate.PaintPicture .ImagePreview, imgW, 0, , , 0, 0, imgW, imgH
            .picTemplate.PaintPicture Picture2, sx, sy
            .picTemplate.Picture = .picTemplate.Image
            For nX = 1 To imgW + 1
                .Picture2.PaintPicture .picTemplate.Picture, 0, 0, , , nX, 0, imgW, imgH
                For nY = 1 To 2000
                    DoEvents
                Next nY
            Next nX
            .picTemplate.Width = imgW
'===============
        Case "up", "up1"
'===============
            .picTemplate.Width = imgW
            .picTemplate.Height = imgH
            .picTemplate.PaintPicture .Picture2.Picture, 0, imgH, , , 0, 0, imgW, imgH
            .picTemplate.PaintPicture ImagePreview, sx, sy
            .picTemplate.Picture = .picTemplate.Image
            For nX = 1 To imgH
                .Picture2.PaintPicture .picTemplate.Picture, 0, 0 + imgH - nX, , , 0, 0, imgW, nX
                For nY = 1 To picTemplate.Height * 4 '2000
                    DoEvents
                Next nY
            Next nX
'=====================================================
        Case "up2" 'fix:20/09/2004
'=====================================================
            .picTemplate.Width = imgW
            .picTemplate.Height = imgH * 2
            picTemplate.PaintPicture .ImagePreview.Picture, 0, imgH, , , 0, 0, imgW, imgH
            .picTemplate.PaintPicture Picture2, sx, sy
            .picTemplate.Picture = .picTemplate.Image
            For nX = 1 To imgH
                .Picture2.PaintPicture .picTemplate.Picture, 0, 0, , , 0, nX, imgW, imgH
                For nY = 1 To 2000
                    DoEvents
                Next nY
            Next nX
            .picTemplate.Height = imgH
'===================================================
        Case "down", "down1"
'===================================================
            .picTemplate.Width = imgW
            .picTemplate.Height = imgH
            Set picTemplate.Picture = picTemplate.Image
            For nX = imgH - 1 To 1 Step -1
                .Picture2.PaintPicture .picTemplate.Image, 0, 0, , , 0, nX, imgW, imgH
                For nY = 1 To picTemplate.Height * 4
                    DoEvents
                Next nY
            Next nX
            Set Picture2.Picture = picTemplate.Image
'===============
        Case "down2"
'===============
            With picTemplate
                .Width = imgW
                .Height = imgH * 2
                .PaintPicture Picture2.Picture, 0, imgH, , , 0, 0, imgW, imgH
                .PaintPicture ImagePreview, sx, sy
                .Picture = .Image
            End With 'picTemplate
            For nX = imgH To 1 Step -1
                Picture2.PaintPicture picTemplate.Picture, 0, 0, , , 0, nX, imgW, imgH
                For nY = 1 To picTemplate.Height * 4
                    DoEvents
                Next nY
            Next nX
' Set Picture2.Picture = Picture2.Image
            picTemplate.Height = imgH
        End Select
'.Picture2.AutoRedraw = True
'.Picture2.Picture = .picHidden1.Picture
    End With
    Set Picture2.Picture = picHidden1.Image
End Sub
Private Sub ShowTransparency(cSrc As PictureBox, _
                             cDest As PictureBox, _
                             ByVal nLevel As Byte)
Dim LrProps    As rBlendProps
Dim LnBlendPtr As Long
    cDest.Cls
    LrProps.tBlendAmount = nLevel
    CopyMemory LnBlendPtr, LrProps, 4
    With cSrc
        AlphaBlend cDest.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, .ScaleWidth, .ScaleHeight, LnBlendPtr
    End With
    cDest.Refresh
End Sub
Private Sub Shrink(ByVal Zoom As Integer, _
                   Optional ByVal strFilter As String)

Dim typRect As TYPERECT
Dim sName   As String
Dim w       As Integer
Dim h       As Integer
    On Error GoTo Pesanku
    sName = GetExtension(m_FileName)
    If InStr(EXT, UCase$(sName)) > 0 Then
        ImagePreview.Cls
    Else
        Exit Sub

    End If
    Picture2.Cls
    picTemplate.Visible = False
    With Picture3
        w = .Width '* 15
        h = .Height ' * 15
    End With 'Picture3
    If w > Zoom Or h > Zoom Then
        If w >= h Then
            ImagePreview.Width = Zoom
            ImagePreview.Height = (h / w) * Zoom
        Else
            ImagePreview.Height = Zoom
            ImagePreview.Width = (w / h) * Zoom
        End If
    Else
        ImagePreview.Width = w
        ImagePreview.Height = h
    End If
    If Picture3.Picture = 0 Then
        Exit Sub

    Else
        ImagePreview.PaintPicture Picture3, 0, 0, ImagePreview.Width, ImagePreview.Height, 0, 0, w, h
    End If
    ImagePreview.Move 0, 0
    With typRect
        .Right = ImagePreview.ScaleWidth ' - 2
        .Top = ImagePreview.ScaleTop '+ 2
        .Left = ImagePreview.ScaleLeft ' + 2   '    .Top = picDest.ScaleWidth
        .Bottom = ImagePreview.ScaleHeight ' - 2
    End With
    Set ImagePreview.Picture = ImagePreview.Image
    imgtempt.Picture = ImagePreview.Image
    Screen.MousePointer = vbDefault
Pesanku:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbCritical + vbOKOnly
    End If
End Sub
Private Sub SliceDown()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    On Error Resume Next
    For PWidth = 40 To picTemplate.ScaleWidth + 40 Step 40
        PHeight = 1
        ms = (TotalDuration / picTemplate.ScaleHeight) / 2
        For I = 1 To picTemplate.ScaleHeight Step 2
            Startplay
            Picture2.AutoRedraw = False
            Picture2.PaintPicture picTemplate.Picture, 0, 0, PWidth, PHeight, 0, 0, PWidth, PHeight, &HCC0020
            PHeight = PHeight + 5
            If PHeight >= Picture2.ScaleHeight + 5 Then
                Exit For
            End If
            Endplay (ms)
        Next I
    Next PWidth
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub SliceLeft()
Dim X    As Integer
Dim Y    As Integer
Dim ms   As Integer
Dim Ytop As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / ImagePreview.ScaleWidth
    For Y = 40 To ImagePreview.ScaleHeight + 40 Step 40 'xHeight
        For X = 1 To ImagePreview.ScaleWidth + 5 Step 5
            Startplay
            Picture2.PaintPicture ImagePreview, Picture2.ScaleWidth - (Picture2.ScaleWidth - ImagePreview.Width) / 2 - X, Ytop, X, Y, ImagePreview.ScaleWidth - X, 0, X, Y, &HCC0020
            Endplay (ms)
        Next X
    Next Y
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub SliceRight()
Dim X  As Integer
Dim Y  As Integer
Dim ms As Integer
    On Error GoTo Pesan
    Picture2.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / picTemplate.ScaleWidth
    For Y = 40 To picTemplate.ScaleHeight + 40 Step 40
        For X = 1 To picTemplate.ScaleWidth + 5 Step 5
            Startplay
            Picture2.PaintPicture picTemplate, 0, 0, X, Y, 0, 0, X, Y, &HCC0020
            Endplay (ms)
        Next X
    Next Y
    Set Picture2.Picture = picTemplate.Image
Pesan:
    PESANUMUM
End Sub
Private Sub SliceUp()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    ms = TotalDuration / ImagePreview.ScaleHeight
    Picture2.Cls
    Picture2.AutoRedraw = False
    For PWidth = (ImagePreview.ScaleWidth / 8) To ImagePreview.ScaleWidth + 40 Step (ImagePreview.ScaleWidth / 8) '20
        PHeight = 1
        For I = 1 To ImagePreview.ScaleHeight
            Startplay
            Picture2.PaintPicture ImagePreview.Picture, 0, ImagePreview.ScaleHeight - PHeight, PWidth, PHeight, 0, ImagePreview.ScaleHeight - PHeight, PWidth, PHeight, &HCC0020
            PHeight = PHeight + 5 '
            If PHeight >= Picture2.Height + 5 Then
                Exit For
            End If
            Endplay (ms)
        Next I
    Next PWidth
    Set Picture2.Picture = ImagePreview.Image
End Sub
Private Sub Slide_BottomLeft()
Dim X  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / ImagePreview.ScaleWidth
    If ImagePreview.ScaleWidth > ImagePreview.ScaleHeight Then
        For X = 1 To ImagePreview.ScaleHeight Step 3
            Startplay
            Picture2.PaintPicture ImagePreview.Picture, -ImagePreview.ScaleWidth + X + (X / 3), ImagePreview.ScaleHeight - X, ImagePreview.ScaleWidth, X, 0, 0, ImagePreview.ScaleWidth, X, &HCC0020
            Endplay (ms)
        Next X
    ElseIf ImagePreview.ScaleHeight > ImagePreview.ScaleWidth Then
        For X = 1 To ImagePreview.ScaleHeight
            Startplay
'                        -(ImagePreview.Width/2+2)atau + 4 bila BorderStyle=1
            Picture2.PaintPicture ImagePreview.Picture, -ImagePreview.ScaleWidth + X - (ImagePreview.Height / 3 + 2), ImagePreview.ScaleHeight - X, ImagePreview.ScaleWidth, X, 0, 0, ImagePreview.ScaleWidth, X, &HCC0020
            Endplay (ms)
        Next X
    End If
    ImagePreview.ScaleHeight = Picture2.ScaleHeight
    Set Picture2.Picture = ImagePreview.Image
End Sub
Private Sub Slide_BottomRight()
Dim X  As Double
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / ImagePreview.ScaleHeight
    If ImagePreview.ScaleWidth > ImagePreview.ScaleHeight Then
        For X = 1 To Picture2.ScaleHeight + 2 Step 3
            Startplay
            Picture2.PaintPicture ImagePreview.Picture, ImagePreview.ScaleWidth - X - (X / 3), ImagePreview.ScaleHeight - X, ImagePreview.ScaleWidth, X, 0, 0, ImagePreview.ScaleWidth, X, &HCC0020
            Endplay (ms)
        Next X
    ElseIf ImagePreview.ScaleHeight > ImagePreview.ScaleWidth Then
        For X = 1 To ImagePreview.ScaleHeight
            Startplay
'
            Picture2.PaintPicture ImagePreview.Picture, ImagePreview.ScaleWidth - X + (ImagePreview.Height / 3 + 2), ImagePreview.ScaleHeight - X, ImagePreview.ScaleWidth, X, 0, 0, ImagePreview.ScaleWidth, X, &HCC0020
            Endplay (ms)
        Next X
    End If
    ImagePreview.ScaleHeight = Picture2.ScaleHeight
    Set Picture2.Picture = ImagePreview.Image
' GoTo Resize
'Resize:
'ImagePreview.Width = UserControl.Width: UserControl.Height = Picture2.Height
End Sub
Private Sub SlideLeft()
Dim X  As Integer
Dim ms As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To Picture2.ScaleWidth Step 1
        Startplay
        Picture2.PaintPicture picTemplate.Picture, -picTemplate.ScaleWidth + X, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub SlideLeft_A()

Dim X  As Integer
Dim ms As Integer
    With Picture2
        .Cls
        .AutoRedraw = False
        .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
    End With 'Picture2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step 3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, -picTemplate.ScaleWidth + X, 0, picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight / 2, 0, 0, picTemplate.ScaleWidth / 2, picTemplate.ScaleHeight / 2, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub SlideRight()
Dim X  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, 0, 0, X, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub SlideRight_RV()
Dim X  As Integer
Dim YX As Integer
Dim ms As Integer
    On Error GoTo Pesan
    Picture2.Cls
    With picEffect
        .Cls
        .AutoRedraw = True
' For YX = 1 To picTemplate.ScaleWidth Step 20
        .PaintPicture picTemplate, picTemplate.ScaleWidth, 0, -picTemplate.ScaleWidth - 1, picTemplate.ScaleHeight, -YX, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
' Next
        Set .Picture = .Image
        .AutoRedraw = False
    End With 'picEffect
    Picture2.Cls
    picEffect.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / picEffect.ScaleWidth
    For X = 1 To picEffect.ScaleWidth Step 3
        Startplay
        Picture2.PaintPicture picEffect.Picture, (picEffect.ScaleWidth + 1) - X, 0, X, picEffect.ScaleHeight, 0, 0, 1 + X, picEffect.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picEffect.Image 'picTemplate.Image
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description
        Exit Sub

    End If
    picEffect.Picture = Nothing
End Sub
Private Sub SlideUp()
Dim X  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To Picture2.ScaleHeight Step 1
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, -picTemplate.ScaleHeight + X, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub

Private Sub Startplay()
    StartTime = timeGetTime()
End Sub
Private Sub StretchClose() 'OKE
Dim Y  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    ms = TotalDuration / picTemplate.ScaleWidth
    For Y = 1 To picTemplate.ScaleHeight / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, Y, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight - Y, &HCC0020
        Picture2.PaintPicture picTemplate.Picture, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, 0, Y, picTemplate.ScaleWidth, picTemplate.ScaleHeight - Y, &HCC0020
        Endplay (ms)
    Next Y
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchDown() 'OKE
Dim Y  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For Y = 1 To picTemplate.ScaleHeight 'Step 3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, &HCC0020
        Endplay ms
    Next Y
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchDown_B() 'OKE
Dim X As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleHeight ' Step 3
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, X, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchLeft() 'OKE
Dim X As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleWidth 'Step 3
        Picture2.PaintPicture picTemplate.Picture, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, -X + picTemplate.ScaleWidth, 0, X, picTemplate.ScaleHeight, &HCC0020
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchLeft_B() 'OKE
Dim X As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleWidth ' Step 3
        Picture2.PaintPicture picTemplate.Picture, 0, 0, X, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchR() 'OKE
Dim X As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleWidth 'Step 3
        Picture2.PaintPicture picTemplate, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, X, picTemplate.ScaleHeight, &HCC0020
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchRight_B() 'OKE
Dim X As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleWidth ' Step 3
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchUp()
Dim X  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For X = 1 To picTemplate.ScaleHeight ' Step 3
        Startplay
        Picture2.PaintPicture picTemplate, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, X, &HCC0020
        Endplay ms
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub StretchUp_B() 'OKE
Dim Y  As Integer
Dim ms As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    For Y = 1 To picTemplate.ScaleHeight Step 3
        Startplay
        Picture2.PaintPicture picTemplate, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, 0, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next Y
    Set Picture2.Picture = picTemplate.Image
End Sub
Public Sub Tile(TileObject As Object, _
                TilePicture As StdPicture)

Dim max_images_width
Dim max_images_height
Dim I           As Integer
Dim ImageTop    As Single
Dim ImageLeft   As Single
Dim ImageWidth  As Single
Dim ImageHeight As Single
Dim PicHolder   As Picture
    On Error GoTo Cancel
    Set PicHolder = TilePicture
    ImageWidth = TileObject.ScaleX(PicHolder.Width, vbHimetric, TileObject.ScaleMode)
    ImageHeight = TileObject.ScaleY(PicHolder.Height, vbHimetric, TileObject.ScaleMode)
    max_images_width = TileObject.ScaleWidth \ ImageWidth
    max_images_height = TileObject.ScaleHeight \ ImageHeight
    TileObject.AutoRedraw = True
    For I = 1 To max_images_height + 1
        For X = 0 To max_images_width
            TileObject.PaintPicture PicHolder, ImageLeft, ImageTop, ImageWidth, ImageHeight
            ImageLeft = ImageLeft + ImageWidth
        Next X
        ImageLeft = 0
        ImageTop = ImageTop + ImageHeight
    Next I
Cancel:
End Sub
Private Sub Translucent(ByVal Opaque As Long)
Dim I  As Integer
Dim ms As Integer
    ms = TotalDuration \ 125
    Picture2.AutoRedraw = True
    ImagePreview.AutoRedraw = True
    Picture2.AutoRedraw = True
    Set Picture2.Picture = Picture2.Image
    For I = 1 To Opaque '255
        Startplay
        ShowTransparency picTemplate, Picture2, I
        Endplay (ms)
    Next I
    Picture2.AutoRedraw = False
    Set Picture2.Picture = Picture2.Image 'picTemplate.Image
End Sub
Private Sub Transparent()
    Picture2.AutoRedraw = True
'Set ImagePreview.Picture = ImagePreview.Image
    MakeTransparent Picture2, picTemplate, Picture2, 50, m_TransValue
    Picture2.AutoRedraw = False
    Set Picture2.Picture = Picture2.Image
End Sub
Public Property Get TransValue() As Integer
    TransValue = m_TransValue
End Property
Public Property Let TransValue(ByVal New_TransValue As Integer)
    m_TransValue = New_TransValue
    If TransValue > 3 Then
        MsgBox "Masukkan angka 1-3.Oke!", vbInformation + vbOKOnly, "PESANKU"
        TransValue = m_def_TransValue
    End If
    PropertyChanged "TransValue"
End Property
Private Sub UserControl_Initialize()
    picHidden1.BackColor = vbBlue
    TotalDuration = 2000
    ms = TotalDuration / picTemplate.ScaleWidth
End Sub
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_TransValue = m_def_TransValue
    m_FileName = m_def_FileName
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        UserControl_Terminate
    End If
End Sub
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_FileName = .ReadProperty("FileName", m_def_FileName)
        m_Value = .ReadProperty("Value", m_def_Value)
        m_TransValue = .ReadProperty("TransValue", m_def_TransValue)
        m_BlockSize = .ReadProperty("BlockSize", m_def_BlockSize)
        m_clearF = .ReadProperty("ClearFirst", m_def_clearF)
    End With 'PropBag
End Sub
Private Sub UserControl_Resize()

    UserControl.Height = (UserControl.ScaleWidth - (UserControl.ScaleWidth / 4)) * 15
    ImagePreview.Width = UserControl.Width / 15
    ImagePreview.Height = UserControl.Height / 15
    Picture2.Width = UserControl.Width / 15
    Picture2.Height = UserControl.Height / 15
    Picture3.Width = UserControl.Width / 15
    Picture3.Height = UserControl.Height / 15
    picTemplate.Width = UserControl.Width / 15
    picTemplate.Height = UserControl.Height / 15
    picEffect.Width = UserControl.Width / 15
    picEffect.Height = UserControl.Height / 15
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Label1.Move 5, Picture2.ScaleHeight - Label1.Height - 5
    Showme
End Sub
Private Sub UserControl_Show()
    Showme
End Sub
Private Sub UserControl_Terminate()
    End
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FileName", m_FileName, m_def_FileName
        .WriteProperty "Value", m_Value, m_def_Value
        .WriteProperty "TransValue", m_TransValue, m_def_TransValue
        .WriteProperty "BlockSize", m_BlockSize, m_def_BlockSize
        .WriteProperty "ClearFirst", m_clearF, m_def_clearF
    End With 'PropBag
End Sub
Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    If Value > 255 Then
        MsgBox "Masukkan angka 1-255.Oke!", vbInformation + vbOKOnly, "PESANKU"
        Value = 255
    End If
    PropertyChanged "value"
End Property
Private Sub SliceVertical1() '#26 'Vertical Allez-Retour une
'By Joko Mulyono \ DANTE '09\09\2003
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    On Error GoTo Pesan
    Picture2.AutoRedraw = False
    For PWidth = 0 To ImagePreview.ScaleWidth Step 80
        PHeight = 1
        ms = (TotalDuration / ImagePreview.ScaleHeight) / 2
        For I = 1 To ImagePreview.ScaleHeight Step 2
            Startplay
            Picture2.PaintPicture picTemplate.Picture, PWidth, 0, 40, PHeight, PWidth, 0, 40, PHeight
            Picture2.PaintPicture picTemplate.Picture, PWidth + 40, picTemplate.ScaleHeight - PHeight, 40, PHeight, PWidth + 40, picTemplate.ScaleHeight - PHeight, 40, PHeight
            PHeight = PHeight + 5
            If PHeight >= Picture2.ScaleHeight + 5 Then
                Exit For
            End If
            Endplay (ms)
        Next I
        If PWidth >= ImagePreview.ScaleWidth + 40 Then
            Exit Sub

        End If
    Next PWidth
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
Pesan:
    If Err.Number <> 0 Then
        If Err.Number = 5 Then
            Resume Next
        End If
    End If
'On Error GoTo 0
End Sub
Private Sub SliceVertical2() '#27 'Vertical Allez-Retour deux
'By Joko Mulyono \ DANTE '09\09\2003
Dim PWidth      As Integer
Dim PHeight     As Integer
Dim I           As Integer
Dim ms          As Integer
Dim Stripewidth As Integer
    Stripewidth = Fix(Picture2.ScaleHeight / 8)
    On Error Resume Next
    Picture2.AutoRedraw = False
    For PWidth = 0 To ImagePreview.ScaleWidth Step Stripewidth * 2
        PHeight = 0
        ms = (TotalDuration / ImagePreview.ScaleHeight) / 2
        For I = 0 To ImagePreview.ScaleHeight Step 2
            Startplay
            Picture2.PaintPicture picTemplate.Picture, PWidth, (-picTemplate.ScaleHeight) + PHeight, Stripewidth, picTemplate.ScaleHeight, PWidth, 0, Stripewidth, picTemplate.ScaleHeight
            Picture2.PaintPicture picTemplate.Picture, PWidth + Stripewidth, (picTemplate.ScaleHeight) - PHeight, Stripewidth, picTemplate.ScaleHeight, PWidth + Stripewidth, 0, Stripewidth, picTemplate.ScaleHeight
            PHeight = PHeight + 5
            If PHeight >= picTemplate.ScaleHeight + 5 Then
                Exit For
            End If
            Endplay (ms)
        Next I
' If PWidth >= ImagePreview.ScaleWidth - 80 Then Exit Sub
    Next PWidth
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub SliceVertical3()
'By Joko Mulyono \ DANTE '09\09\2003
Dim PWidth  As Integer
Dim PHeight As Integer

Dim I       As Integer
Dim N       As Integer
Dim ms      As Integer
    On Error Resume Next
    Picture2.AutoRedraw = False
    For PWidth = 0 To ImagePreview.ScaleWidth Step 80
        PHeight = 0
        ms = (TotalDuration / ImagePreview.ScaleHeight) / 2
        For I = 1 To ImagePreview.ScaleHeight Step 2
            Startplay
'Picture2.AutoRedraw = False
            Picture2.PaintPicture picTemplate.Picture, PWidth, -picTemplate.ScaleHeight + 5 + PHeight, 40, picTemplate.ScaleHeight, PWidth, 0, 40, picTemplate.ScaleHeight
            PHeight = PHeight + 5
            If PHeight >= picTemplate.ScaleHeight Then
                Exit For
            End If
            Endplay (ms)
        Next I
        PHeight = 0
        For N = 1 To ImagePreview.ScaleHeight Step 2
            Startplay
            Picture2.AutoRedraw = False
            Picture2.PaintPicture picTemplate.Picture, PWidth + 40, picTemplate.ScaleHeight - 5 - PHeight, 40, picTemplate.ScaleHeight, PWidth + 40, 0, 40, picTemplate.ScaleHeight
            PHeight = PHeight + 5
            If PHeight >= picTemplate.ScaleHeight Then
                Exit For
            End If
            Endplay (ms)
        Next N
'If PWidth >= picTemplate.ScaleWidth - 80 Then Exit Sub
    Next PWidth
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
    On Error GoTo 0
End Sub
Private Sub WaveDown()
Dim Y  As Integer
Dim ms As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
'If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For Y = 1 To picTemplate.ScaleHeight Step 3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, picTemplate.ScaleHeight - Y, picTemplate.ScaleWidth, Y, &HCC0020
        Endplay (ms)
    Next Y
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub WaveLeft()
' ^pic Resource   ^ picDest
Dim X As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step 3
        Picture2.PaintPicture picTemplate.Picture, -picTemplate.ScaleWidth + X, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, X, picTemplate.ScaleHeight, &HCC0020
    Next X
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub WaveRight()
Dim X As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step 3
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, picTemplate.ScaleWidth, picTemplate.ScaleHeight, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, &HCC0020
    Next X
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub WaveUp()
Dim Y  As Integer
Dim ms As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
'If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For Y = 1 To picTemplate.ScaleHeight Step 3
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, -picTemplate.ScaleHeight + Y, picTemplate.ScaleWidth, picTemplate.ScaleHeight, 0, 0, picTemplate.ScaleWidth, Y, &HCC0020
        Endplay (ms)
    Next Y
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesCenter()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    If picTemplate.ScaleWidth > picTemplate.ScaleHeight Then
        PWidth = picTemplate.ScaleWidth - picTemplate.ScaleHeight
        PHeight = 1
    ElseIf picTemplate.ScaleWidth < picTemplate.ScaleHeight Then
        PWidth = 1
        PHeight = picTemplate.ScaleHeight - picTemplate.ScaleWidth
    Else
        PWidth = 1
        PHeight = 1
    End If
    ms = TotalDuration / (picTemplate.ScaleWidth - PWidth)
    For I = 1 To picTemplate.ScaleWidth - PWidth
        Startplay
        Picture2.PaintPicture picTemplate.Picture, Int((picTemplate.ScaleWidth - PWidth) / 2), Int((picTemplate.ScaleHeight - PHeight) / 2), PWidth, PHeight, Int((picTemplate.ScaleWidth - PWidth) / 2), Int((picTemplate.ScaleHeight - PHeight) / 2), PWidth, PHeight, &HCC0020
        PWidth = PWidth + 5
        PHeight = Height + 5
        If PWidth > Picture2.Width + 5 Then
            Exit For
        End If
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesClose()
Dim X  As Integer
Dim ms As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, X, picTemplate.ScaleHeight, 0, 0, X, picTemplate.ScaleHeight, &HCC0020
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesCloseV()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    PWidth = picTemplate.ScaleWidth
    PHeight = 1
    ms = (TotalDuration / picTemplate.ScaleHeight) / 2
    For I = 1 To picTemplate.ScaleHeight / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, (picTemplate.ScaleHeight - PHeight) / 2, PWidth, PHeight, 0, (picTemplate.ScaleHeight - PHeight) / 2, PWidth, PHeight, &HCC0020
        PHeight = PHeight + 2
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesLeft()
Dim X  As Integer
Dim ms As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth Step 5
        Startplay
        Picture2.PaintPicture picTemplate.Picture, 0, 0, X, picTemplate.ScaleHeight, 0, 0, X, picTemplate.ScaleHeight, &HCC0020
        Endplay (ms)
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesOpenH()
Dim PWidth  As Integer
Dim PHeight As Integer
Dim I       As Integer
Dim ms      As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    Set picTemplate.Picture = picTemplate.Image
    PWidth = 1
    PHeight = picTemplate.ScaleHeight
    ms = TotalDuration / (picTemplate.ScaleWidth / 2)
    For I = 1 To picTemplate.ScaleWidth / 2
        Startplay
        Picture2.PaintPicture picTemplate.Picture, (picTemplate.ScaleWidth - PWidth) / 2, 0, PWidth, PHeight, (picTemplate.ScaleWidth - PWidth) / 2, 0, PWidth, PHeight, &HCC0020
        PWidth = PWidth + 2
        Endplay (ms)
    Next I
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub WipesRights()
Dim X  As Integer
Dim ms As Integer
    Picture2.AutoRedraw = False
    Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
    Set picTemplate.Picture = picTemplate.Image
    For X = 1 To picTemplate.ScaleWidth
        Startplay
        Picture2.PaintPicture picTemplate.Picture, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, picTemplate.ScaleWidth - X, 0, X, picTemplate.ScaleHeight, &HCC0020
        Endplay ms
    Next X
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub Zoom_LeftBottom()
Dim w As Integer
Dim h As Integer
Dim X As Integer

    Picture2.Cls
    Picture2.AutoRedraw = False
    With picTemplate
        Set .Picture = .Image
        w = .Width
        h = .Height
        For X = 1 To .ScaleWidth + 3 Step 3
            .Width = X - 1
'fix:26/05/2005 'added Int
            .Height = Int((h / w) * X)  'tray to remove this line
            Picture2.PaintPicture picTemplate, 0, Picture2.ScaleHeight - (X - (X / 4)), X, .ScaleHeight
        Next X
    End With 'picTemplate
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub Zoom_RightBottom()
Dim w As Integer
Dim h As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    With picTemplate
        Set .Picture = .Image
        w = .Width
        h = .Height
        For X = 1 To .ScaleWidth Step 3
'fix:26/05/2005 added int and + 2
            .Width = X + 2
            .Height = Int((h / w) * X) + 2
            Picture2.PaintPicture picTemplate, Picture2.ScaleWidth - X, Picture2.ScaleHeight - (h / w) * X, X, .ScaleHeight
        Next X
    End With 'picTemplate
    Picture2.AutoRedraw = False
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub Zoom_UpLeft()
Dim w As Integer
Dim h As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    With picTemplate
        Set .Picture = .Image
        w = .Width
        h = .Height
        For X = 1 To .ScaleWidth Step 3
'fix:26/05/2005 added int and + 2
            .Width = X + 2
            .Height = Int((h / w) * X) + 2
            Picture2.PaintPicture picTemplate, 0, 0, .ScaleWidth, .ScaleHeight
        Next X
    End With 'picTemplate
    Picture2.AutoRedraw = False
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub Zoom_UpRight()
Dim w As Integer
Dim h As Integer
    Picture2.Cls
    Picture2.AutoRedraw = False
    With picTemplate
        Set .Picture = .Image
        w = .Width
        h = .Height
        For X = 1 To .ScaleWidth Step 3
'fix:26/05/2005 added Int and + 2
            .Width = X + 2
            .Height = Int((h / w) * X) + 2
            Picture2.PaintPicture picTemplate, Picture2.ScaleWidth - X, 0, X, .ScaleHeight
        Next X
    End With 'picTemplate
    Picture2.AutoRedraw = False
    Picture2.Picture = picTemplate.Image
End Sub
Private Sub ZoomIn() 'Not perfect
Dim I As Integer
Dim T As Integer
    For I = 0 To 750 Step 10
        With Picture2
            Picture2.Cls
            Picture2.AutoRedraw = True
            Shrink (800 - I)
            Set .Picture = ImagePreview.Image
            .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
            .AutoRedraw = False
        End With 'Picture2
        If I = 750 Then
            For T = 10 To (UserControl.Width / 15) + 20 Step 10
                Set Picture2.Picture = ImagePreview.Image
                Picture2.Move (UserControl.ScaleWidth - Picture2.Width) / 2, (UserControl.ScaleHeight - Picture2.Height) / 2
                Shrink (T)
            Next T
        End If
    Next I
End Sub
Private Sub ZoomLeft_B()
Dim X As Integer

    Picture2.Cls
    Picture2.AutoRedraw = False
    With picTemplate
        Set .Picture = .Image

        For X = 1 To .ScaleWidth + 3 Step 3
            .Width = X - 1
            Picture2.PaintPicture picTemplate, 0, Picture2.ScaleHeight - (X - (X / 4)), X, .ScaleHeight
        Next X
    End With 'picTemplate
    Picture2.AutoRedraw = True
    Set Picture2.Picture = picTemplate.Image
End Sub
Private Sub ZoomOut() 'Not perfect
Dim I As Integer
    For I = 10 To (UserControl.Width / 15) + 20 Step 10
        Picture2.Cls
        Shrink (I)
        With Picture2
            .Cls
            .AutoRedraw = True
            Set .Picture = ImagePreview.Image
            .Move (UserControl.ScaleWidth - .Width) / 2, (UserControl.ScaleHeight - .Height) / 2
            .AutoRedraw = False
        End With 'Picture2
    Next I
End Sub

'Note:
'SliceVertical3    :Fix 30/05/2005:10:35 AM
'SliceHorizontal1  :fix 31/05/2005:12:28 AM
'SliceVertical2    :Fix 31/05/2005:12:38 AM
'SliceVertical1    :Fix 31/05/2005:12:47 AM

