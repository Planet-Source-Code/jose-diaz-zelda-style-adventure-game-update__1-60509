Attribute VB_Name = "GAMEDATA"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Key values
Global Const KEY_ESCAPE = 1
Global Const KEY_1 = 2
Global Const KEY_2 = 3
Global Const KEY_3 = 4
Global Const KEY_4 = 5
Global Const KEY_5 = 6
Global Const KEY_6 = 7
Global Const KEY_7 = 8
Global Const KEY_8 = 9
Global Const KEY_9 = 10
Global Const KEY_0 = 11
Global Const KEY_MINUS = 12
Global Const KEY_EQUALS = 13
Global Const KEY_BACKSPACE = 14
Global Const KEY_TAB = 15
Global Const KEY_Q = 16
Global Const KEY_W = 17
Global Const KEY_E = 18
Global Const KEY_R = 19
Global Const KEY_T = 20
Global Const KEY_Y = 21
Global Const KEY_U = 22
Global Const KEY_I = 23
Global Const KEY_O = 24
Global Const KEY_P = 25
Global Const KEY_LBRACKET = 26
Global Const KEY_RBRACKET = 27
Global Const KEY_RETURN = 28
Global Const KEY_LCONTROL = 29
Global Const KEY_A = 30
Global Const KEY_S = 31
Global Const KEY_D = 32
Global Const KEY_F = 33
Global Const KEY_G = 34
Global Const KEY_H = 35
Global Const KEY_J = 36
Global Const KEY_K = 37
Global Const KEY_L = 38
Global Const KEY_SEMICOLON = 39
Global Const KEY_APOSTROPHE = 40
Global Const KEY_GRAVE = 41
Global Const KEY_LSHIFT = 42
Global Const KEY_BACKSLASH = 43
Global Const KEY_Z = 44
Global Const KEY_X = 45
Global Const KEY_C = 46
Global Const KEY_V = 47
Global Const KEY_B = 48
Global Const KEY_N = 49
Global Const KEY_M = 50
Global Const KEY_COMMA = 51
Global Const KEY_PERIOD = 52
Global Const KEY_SLASH = 53
Global Const KEY_RSHIFT = 54
Global Const KEY_MULTIPLY = 55
Global Const KEY_LALT = 56
Global Const KEY_SPACE = 57
Global Const KEY_CAPSLOCK = 58
Global Const KEY_F1 = 59
Global Const KEY_F2 = 60
Global Const KEY_F3 = 61
Global Const KEY_F4 = 62
Global Const KEY_F5 = 63
Global Const KEY_F6 = 64
Global Const KEY_F7 = 65
Global Const KEY_F8 = 66
Global Const KEY_F9 = 67
Global Const KEY_F10 = 68
Global Const KEY_NUMLOCK = 69
Global Const KEY_SCROLL = 70
Global Const KEY_NUMPAD7 = 71
Global Const KEY_NUMPAD8 = 72
Global Const KEY_NUMPAD9 = 73
Global Const KEY_SUBTRACT = 74
Global Const KEY_NUMPAD4 = 75
Global Const KEY_NUMPAD5 = 76
Global Const KEY_NUMPAD6 = 77
Global Const KEY_ADD = 78
Global Const KEY_NUMPAD1 = 79
Global Const KEY_NUMPAD2 = 80
Global Const KEY_NUMPAD3 = 81
Global Const KEY_NUMPAD0 = 82
Global Const KEY_DECIMAL = 83
Global Const KEY_F11 = 87
Global Const KEY_F12 = 88
Global Const KEY_NUMPADENTER = 156
Global Const KEY_RCONTROL = 157
Global Const KEY_DIVIDE = 181
Global Const KEY_RALT = 184
Global Const KEY_HOME = 199
Global Const KEY_UP = 200
Global Const KEY_PAGEUP = 201
Global Const KEY_LEFT = 203
Global Const KEY_RIGHT = 205
Global Const KEY_END = 207
Global Const KEY_DOWN = 208
Global Const KEY_PAGEDOWN = 209
Global Const KEY_INSERT = 210
Global Const KEY_DELETE = 211

'font variables
Global Const TEXT_TRANSPARENT = 1       'Textout constants
Global Const TEXT_OPAQUE = 2
Global Const FW_NORMAL = 400            'Font constants
Global Const DEFAULT_CHARSET = 1
Global Const OUT_TT_ONLY_PRECIS = 7
Global Const CLIP_DEFAULT_PRECIS = 0
Global Const CLIP_LH_ANGLES = &H10
Global Const PROOF_QUALITY = 2
Global Const TRUETYPE_FONTTYPE = &H4
Public lngDC As Long
Public lngNewFont As Long
Public lngOldFont As Long

'sound vars
Public Loader As DirectMusicLoader

' this controls the music
Public Performance As DirectMusicPerformance

' this stores the music in memory
Public Segment As DirectMusicSegment

Public Type COORD
    X As Integer
    Y As Integer
End Type

Public Type gmPicture
    Pic As DirectDrawSurface7
    rRect As RECT
    Width As Integer
    Height As Integer
End Type

Public Type NormSurface
    rRect As RECT
    Surface As DirectDrawSurface7
End Type


Public rScreen As RECT
Public surTiles As gmPicture
Public surTiles2 As gmPicture
Public surTiles3 As gmPicture
Public surMessageBox As gmPicture
Public surChars As gmPicture
Public DBuffer2 As NormSurface
Public DBuffer3 As NormSurface
Public surScroll As NormSurface
'All the properties of the player thus far
Public Type Char
    Xpos As Integer
    Ypos As Integer
    TR As Integer
    TC As Integer
    TopLeft As COORD
    TopRight As COORD
    BottomL As COORD
    BottomR As COORD
    OldXrow As Integer
    OldYcol As Integer
    Xrow As Integer
    Ycol As Integer
    Moving As Boolean
    CharIDX As Integer
    CharIDY As Integer
    Frame As Integer
    Pushing As Boolean
    Attacking As Boolean
    AttackFrame As Integer
    OldIDX As Integer
    OldIDY As Integer
End Type
'tell me which way the player was facing when he pressed an action key
Public Enum Action
    ActionUp
    ActionDown
    ActionLeft
    ActionRight
End Enum

Public DirAction As Action

Public Type EventsInfo
    Type As Integer
    Message As String
End Type
'holds the info that i need in order to know which map to load and where to send the char to
Public Type LinkInfo
    MapName As String
    GotoCoord As COORD
    PlayerDir As Integer
End Type
'will be used later when i implement the sword
Public Type Sword
    IDX As Integer
    IDY As Integer
    Damage As Integer
    Frame As Integer
    FrameCDir As Integer
End Type

Public Type layer
    Used As Byte
    TileX As Integer
    TileY As Integer
End Type
'hold the properties each tile in the map has
Public Type MapData
    layer(0 To 2) As layer
    Blocked As Byte
    Animated As Byte
    Linked As Byte
    LinkInfo As LinkInfo
    Event As EventsInfo
End Type

Public FlipTime As Long
Public Player As Char

Public MapRow As Integer
Public MapCol As Integer
Public curScrollRow As Integer
Public MaxRow As Integer
Public MaxCol As Integer
Public Map() As MapData
Public TownMap() As MapData
Public Sword As Sword
Public AnimatedWaterTiles() As COORD
Public AnimatedGrassTiles() As COORD
Public PlaceScreenX As Integer
Public PlaceScreenY As Integer
Public ScreenX As Integer
Public ScreenY As Integer
Public WaterUp As Boolean
Public GrassUp As Boolean
Public WaterFrameCount As Integer
Public GrassFrameCount As Integer
Public opentime As Long
Public AnimTime As Long
Public AnimGTime As Long
Public MapStartX As Integer
Public MapStartY As Integer
Public PlayerStartDir As Integer
Public ScreenAdjX As Integer
Public ScreenAdjY As Integer
Public NumLetter As Integer
Public DoneWithMessage As Boolean
Public AbletoPress As Boolean
Public CurEvent As Boolean
Public finAttack As Boolean
Public scrRect As RECT
