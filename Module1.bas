Attribute VB_Name = "Module1"
'<:-) :WARNING: All variables must now be declared.
'<:-) Run code using [Ctrl]+[F5] to find undeclared variables that Code Fixer misses.
Option Explicit
#If Win32 Then
'<:-) :SUGGESTION: Is 16-Bit support necessary?
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, _
                                                                             ByVal uFlags As Long) As Long
'<:-) :WARNING: Scope Changed to Private
'<:-) :WARNING: Scope Changed to Public
#Else
Private Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName As Any, _
                                                          ByVal wFlags As Integer) As Integer
'<:-) :WARNING: Scope Changed to Private
'<:-) :WARNING: Scope Changed to Public
#End If
Private Const SND_ASYNC        As Long = &H1    'Play asynchronously
'<:-) :WARNING: Scope Changed to Private
'<:-) :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
'<:-) :UPDATED: Module Level 'Global' to 'Public'
Private Const SND_NODEFAULT    As Long = &H2    'Don't use default sound
'<:-) :WARNING: Scope Changed to Private
'<:-) :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
'<:-) :UPDATED: Module Level 'Global' to 'Public'
Private Const SND_MEMORY       As Long = &H4    'lpszSoundName points to a memory file
'<:-) :WARNING: Scope Changed to Private
'<:-) :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
'<:-) :UPDATED: Module Level 'Global' to 'Public'
''Private Const SND_LOOP         As Long = &H8
'<:-) :WARNING: Unused Const 'SND_LOOP'
'<:-) :WARNING: Scope Changed to Private
'<:-) :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
'<:-) :UPDATED: Module Level 'Global' to 'Public'
Private Const SND_NOSTOP       As Long = &H10
'<:-) :WARNING: Scope Changed to Private
'<:-) :UPDATED: UnTyped Const with Hex (&H) value changed to As Long
'<:-) :UPDATED: Module Level 'Global' to 'Public'
Private SoundBuffer            As String
'<:-) :WARNING: Scope Changed to Private
''
'''<:-) :UPDATED: Module Level 'Global' to 'Public'
''Private Sub BeginPlaySound(ByVal ResourceId As Integer)
''
'''<:-) :WARNING: Unused Sub 'BeginPlaySound'
'''<:-) :WARNING: Scope Changed to Private
'''Dim Ret As Variant
'''<:-) :WARNING: Variable is assigned a Return value from an API Function call but never used in code.
'''<:-) :SUGGESTION: The Function may be setting one of its parameters for use in code.
'''<:-) :WARNING: Code Fixer has replaced assignment with a direct call to the Function
''#If Win32 Then
'''<:-) :SUGGESTION: Is 16-Bit support necessary?
''' Important: The returned string is converted to Unicode
''SoundBuffer = StrConv(LoadResData(ResourceId, "ATM_SOUND"), vbUnicode)
''#Else
''SoundBuffer = LoadResData(ResourceId, "ATM_SOUND")
''#End If
''sndPlaySound SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_NOSTOP
'''<:-) :WARNING: assigned only variable 'Ret' removed.
''' Important: This function is neccessary for playing sound asynchronously
''DoEvents
''End Sub
''
''
''Private Sub EndPlaySound()
''
'''<:-) :WARNING: Unused Sub 'EndPlaySound'
'''<:-) :WARNING: Scope Changed to Private
'''Dim Ret As Variant
'''<:-) :WARNING: Variable is assigned a Return value from an API Function call but never used in code.
'''<:-) :SUGGESTION: The Function may be setting one of its parameters for use in code.
'''<:-) :WARNING: Code Fixer has replaced assignment with a direct call to the Function
''sndPlaySound 0&, 0&
'''<:-) :WARNING: assigned only variable 'Ret' removed.
''End Sub
''
':)Code Fixer V3.0.9 (5/31/2005 8:00:00 AM) 75 + 0 = 75 Lines Thanks Ulli for inspiration and lots of code.


