VERSION 5.00
Begin VB.Form setting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ControlBox      =   0   'False
   Icon            =   "setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   243
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   6360
      Picture         =   "setting.frx":030A
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   2100
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   105
         Index           =   2
         Left            =   1080
         Picture         =   "setting.frx":E8FC
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   105
         Index           =   1
         Left            =   1080
         Picture         =   "setting.frx":EA02
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   105
         Index           =   2
         Left            =   0
         Picture         =   "setting.frx":EB08
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   105
         Index           =   1
         Left            =   0
         Picture         =   "setting.frx":EC0E
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000010&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   1
      Top             =   120
      Width           =   3405
      Begin VB.PictureBox Picture8 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "setting.frx":ED14
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   3600
         Width           =   480
         Begin VB.PictureBox Picture9 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   105
            Index           =   0
            Left            =   330
            Picture         =   "setting.frx":F356
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   10
            Top             =   120
            Width           =   135
         End
         Begin VB.PictureBox Picture10 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   105
            Index           =   0
            Left            =   330
            Picture         =   "setting.frx":F45C
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   9
            Top             =   15
            Width           =   135
         End
         Begin VB.Label lblDelay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000016&
         Height          =   2985
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H8000000D&
         Height          =   450
         Left            =   2040
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   3
         Top             =   3480
         Width           =   1215
      End
      Begin VB.PictureBox picAbout 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   720
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   2
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Select filename:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4320
      Width           =   1380
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   360
      Pattern         =   "*.jpg;*.gif;*.bmp;*.wmf"
      ReadOnly        =   0   'False
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Delay        As Integer
Private TheGreater   As Integer
Private TheText      As String
Private TheGrestest  As Integer
Private oldfolder    As String
Private Const LB_GETSELCOUNT             As Long = &H190
Private Const LB_SETHORIZONTALEXTENT     As Long = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()

Tile Me, Picture1.Picture
PaintControl picAbout, HalfRaised, &H80000010, vbWhite, "Cancel", False
PaintControl Picture5, HalfRaised, RGB(132, 137, 141), vbWhite, "", False
PaintControl Picture6, HalfRaised, &H80000010, vbWhite, "Browse", False
lblDelay.Caption = GetSetting("file", "settings", "time", vbNullString)
LoadDIR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub picAbout_Click()
SaveSetting "setting", "directory", "mDir", oldfolder
Unload Me
End Sub

Private Sub picAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PaintControl picAbout, Bump, vbBlack, vbWhite, "Browse", False
End Sub

Private Sub picAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PaintControl picAbout, HalfRaised, &H80000010, vbWhite, "Cancel", False
End Sub

Private Sub Picture10_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture10(0).Picture = Picture10(2).Picture
If Delay >= 8 Then
   Exit Sub
End If
Delay = Delay + 1
lblDelay.Caption = Delay
SaveSetting "file", "settings", "time", CStr(Delay)
End Sub

Private Sub Picture10_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture10(0).Picture = Picture10(1).Picture
End Sub

Private Sub Picture6_Click()
Dim i As String
i = GetSetting("setting", "directory", "mDir", vbNullString)
oldfolder = i
FolderPath = GET_DIRECTORY(Me)
File1.Path = FolderPath
List1.Clear
LoadListDir
End Sub
Private Sub LoadListDir()

Dim i           As Integer
Static X        As Long

    TheGreater = Len(List1.Text)
    List1.Clear
    For i = 1 To File1.ListCount '- 1
        File1.ListIndex = i - 1
        List1.AddItem File1.FileName
        List1.ListIndex = i - 1
        If Len(List1.Text) > TheGreater Then
            TheGreater = Len(List1.Text)
            TheText = List1.Text
            TheGrestest = List1.ListIndex
        End If
        
    Next i
    If X < TextWidth(TheText & "  ") Then
        X = TextWidth(TheText & "  ")
        If ScaleMode = vbTwips Then
            X = X / Screen.TwipsPerPixelX
        End If
        'Horizontal scroll bar
        SendMessageByNum List1.hWnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
    If List1.ListCount <> 0 Then
        List1.ListIndex = TheGrestest
        
    End If
End Sub
Private Sub LoadDIR()
Dim i As String
i = GetSetting("setting", "directory", "mDir", vbNullString)

File1.Path = i
FolderPath = i
List1.Clear
LoadListDir
End Sub
Private Sub Tile(TileObject As Object, _
                 TilePicture As StdPicture)
Dim max_images_width
Dim max_images_height
Dim i           As Integer
Dim ImageTop    As Single
Dim ImageLeft   As Single
Dim ImageWidth  As Single
Dim ImageHeight As Single
Dim PicHolder   As Picture
Dim X As Integer
    On Error GoTo Cancel
    Set PicHolder = TilePicture
    ImageWidth = TileObject.ScaleX(PicHolder.Width, vbHimetric, TileObject.ScaleMode)
    ImageHeight = TileObject.ScaleY(PicHolder.Height, vbHimetric, TileObject.ScaleMode)
    max_images_width = TileObject.ScaleWidth \ ImageWidth
    max_images_height = TileObject.ScaleHeight \ ImageHeight
    TileObject.AutoRedraw = True
    For i = 1 To max_images_height + 1
        For X = 0 To max_images_width
            TileObject.PaintPicture PicHolder, ImageLeft, ImageTop, ImageWidth, ImageHeight
            ImageLeft = ImageLeft + ImageWidth
        Next X
        ImageLeft = 0
        ImageTop = ImageTop + ImageHeight
    Next i
Cancel:
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PaintControl Picture6, Bump, vbBlack, vbWhite, "Browse", True

End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PaintControl Picture6, HalfRaised, &H80000010, vbWhite, "Browse", False
End Sub

Private Sub Picture9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture9(0).Picture = Picture9(2).Picture
If Delay <= 1 Then
   Exit Sub
End If
Delay = Delay - 1
lblDelay.Caption = Delay
SaveSetting "file", "settings", "time", CStr(Delay)
End Sub

Private Sub Picture9_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture9(0).Picture = Picture9(1).Picture
End Sub

