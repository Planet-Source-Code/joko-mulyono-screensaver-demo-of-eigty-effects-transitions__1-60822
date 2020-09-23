VERSION 5.00
Begin VB.Form Slide 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   10275
   ClientLeft      =   135
   ClientTop       =   510
   ClientWidth     =   11520
   ControlBox      =   0   'False
   Icon            =   "Slide.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ScreenSaver.ImageEffect ImageEffect1 
      Height          =   6945
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12250
      BlockSize       =   0
   End
   Begin VB.PictureBox picture_loaded 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   3840
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   4
      Top             =   1560
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   855
      Left            =   5400
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1980
      Left            =   7080
      Pattern         =   "*.bmp;*.jpg;*.gif"
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PATH"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   9000
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EFFECT NAME"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   9360
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   2
      Left            =   7680
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   0
      Left            =   -90
      TabIndex        =   1
      Top             =   10710
      Visible         =   0   'False
      Width           =   4845
   End
End
Attribute VB_Name = "Slide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private nEffect         As Integer
Private PicNum          As Integer
Private HC              As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub Form_Activate()
Dim timeset As Long

   
    If File1.ListCount > 100 Then
        timeset = GetSetting("file", "settings", "time", "1000")
        Timer1.Interval = timeset * 1000
    Else
        Timer1.Interval = 5000
    End If
    Timer1.Enabled = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'32=SPACEBAR
    If KeyAscii = 32 Then
        setting.Show
    End If
    HC = ShowCursor(True)
    Timer1.Enabled = False
'End
End Sub
Private Sub Form_Load()
Dim i As String
i = GetSetting("setting", "directory", "mDir", vbNullString)
    If i <> vbNullString Then
         File1.Path = i
         FolderPath = i
    Else
         File1.Path = "C:\WINDOWS"
         FolderPath = "C:\WINDOWS"
    End If
'------------------------------------------------- new line
    Picture2.Move (Screen.Width / 15 - Picture2.Width) / 2, (Screen.Height / 15 - Picture2.Height) / 2

    With ImageEffect1
        .ClearFirst = True
        .Width = 800
        .Height = 600
        .BlockSize = 5
        .Move (Screen.Width / 15 - .Width) / 2, (Screen.Height / 15 - .Height) / 2
        Label3.Move .Left, .Top + .Height + 10
        Label2.Move .Left, .Top + .Height + 30
    End With 'ImageEffect1
    If App.PrevInstance Then

        End
    End If
    Select Case LCase$(Left$(Command, 2))
    Case "/p"
        Me.Hide
        End
        Exit Sub
    Case "/s"

'Proceed
    Case Else
'Show settings
        Me.Hide
        setting.Show
        Exit Sub
    End Select
    HC = ShowCursor(False)
End Sub
Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    HC = ShowCursor(True)
    End
End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
Static count As Integer

    count = count + 1
    If count > 5 Then
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    HC = ShowCursor(True)
    End
End Sub

Private Sub Timer1_Timer()

    Label3.Caption = FolderPath & "\" & File1.List(PicNum)
    Randomize
    nEffect = Rnd * (85 + 1)
    Label2.Caption = ImageEffect1.GetEffectName(nEffect)
    ImageEffect1.FileName = FolderPath & "\" & File1.List(PicNum)
    ImageEffect1.LoadGambar FolderPath & "\" & File1.List(PicNum), nEffect
    If PicNum = File1.ListCount - 1 Then
        PicNum = 0
    Else
        PicNum = PicNum + 1
    End If
End Sub


