VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DX7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dx As DirectX7
Public dd As DirectDraw7
Public dxWhndl As Long
Public primary As DirectDrawSurface7
Public backbuffer As DirectDrawSurface7
Private ddsdBackbuffer As DDSURFACEDESC2
Private rBack As RECT
Public ScreenWidth As Integer
Public ScreenHeight As Integer

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BitmapStruct
    FileHeader As BITMAPFILEHEADER
    InfoHeader As BITMAPINFOHEADER
End Type
Dim ddsd As DDSURFACEDESC2
Dim BITMAP As BitmapStruct

'DIRECT INPUT VARS
Dim di As DirectInput
Dim diDEV As DirectInputDevice
Dim diState As DIKEYBOARDSTATE

'Direct Sound Vars
Public ds As DirectSound
Public DSBuffer As DirectSoundBuffer



Public Function InitDX(lngWindowHandle As Long, lWidth As Long, lHeight As Long, colordepth As Integer) As Boolean
    
    On Error GoTo ErrorHand
    
    Dim ddsd As DDSURFACEDESC2
    Dim dds As DDSCAPS2
    
    Set backbuffer = Nothing
    Set primary = Nothing
    Set dx = New DirectX7
    
    Set dd = dx.DirectDrawCreate("")
    
    dxWhndl = lngWindowHandle
    
    dd.SetCooperativeLevel lngWindowHandle, DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX
    'Set the display mode
    dd.SetDisplayMode lWidth, lHeight, colordepth, 0, DDSDM_DEFAULT
    
    ScreenWidth = lWidth
    ScreenHeight = lHeight
    rScreen.Bottom = lHeight
    rScreen.Right = lWidth
    'Create the primary surface
    ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd.lWidth = ScreenWidth
    ddsd.lHeight = ScreenHeight
    ddsd.lBackBufferCount = 2
    Set primary = dd.CreateSurface(ddsd)
    
    dds.lCaps = DDSCAPS_BACKBUFFER
    
    'Create the backbuffer
    rBack.Bottom = ScreenHeight
    rBack.Right = ScreenWidth
    rBack.Top = 0
    rBack.Left = 0
    Set backbuffer = primary.GetAttachedSurface(dds)
    backbuffer.GetSurfaceDesc ddsdBackbuffer
    '**************INIT Direct Input***********************
    'Create the direct input object
    Set di = dx.DirectInputCreate()
        
    'Get the keyboard
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    
    'Get input nonexclusively, only when in foreground mode
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    diDEV.Acquire
    
   
    Set Loader = dx.DirectMusicLoaderCreate
    ' create the performance
    Set Performance = dx.DirectMusicPerformanceCreate
    
    ' start up the performance, telling DirectMusic
    ' the handle of the form
    Performance.Init Nothing, lngWindowHandle
    ' set the port (-1 lets DirectMusic to choose
    ' the port itself - less work for us)
    Performance.SetPort -1, 1
    ' tell DirectMusic to do all the sound downloading
    ' stuff itself because we can't be bothered to do it
    Performance.SetMasterAutoDownload True

    
    
    InitDX = True
    
    Exit Function
ErrorHand:
    InitDX = False
   
    MsgBox "Error DD Init: " & Err.Number & " / " & Err.Description, vbCritical
    
End Function

Public Function LoadDDBitMap(FileName As String, MakeTrans As Boolean, Optional TransColorHigh As Long, Optional TransColorLow As Long) As DirectDrawSurface7

    Dim ddsd As DDSURFACEDESC2
    Dim TransKey As DDCOLORKEY
    
    ddsd.lFlags = DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ddsd.lWidth = GetBitmapInfo(FileName, False, True)
    ddsd.lHeight = GetBitmapInfo(FileName, True, False)
    
    'load the bitmap from the file
    Set LoadDDBitMap = dd.CreateSurfaceFromFile(FileName, ddsd)
   
    'Set the transparent color for the bitmap
    If MakeTrans = True Then
        TransKey.high = TransColorHigh
        TransKey.low = TransColorLow

        LoadDDBitMap.SetColorKey DDCKEY_SRCBLT, TransKey
    End If
    
End Function

Public Function MakeNormSurface(intWidth As Long, intHeight As Long, MakeTrans As Boolean, Optional TransColorHigh As Long, Optional TransColorLow As Long) As DirectDrawSurface7
    Dim ddsd As DDSURFACEDESC2
    Dim TransKey As DDCOLORKEY
    
    ddsd.lFlags = DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    ddsd.lWidth = intWidth
    ddsd.lHeight = intHeight
    
    
    Set MakeNormSurface = dd.CreateSurface(ddsd)
   
    'Set the transparent color for the bitmap
    If MakeTrans = True Then
        TransKey.high = TransColorHigh
        TransKey.low = TransColorLow
        
        MakeNormSurface.SetColorKey DDCKEY_SRCBLT, TransKey
    End If
End Function

Public Sub BltSurface(Destination As DirectDrawSurface7, Source As DirectDrawSurface7, Xpos As Integer, Ypos As Integer, Top, Left, Width, Height)
    Dim rRect As RECT
    
    rRect.Top = Top
    rRect.Left = Left
    rRect.Bottom = Top + Height
    rRect.Right = Left + Width
    
    Destination.BltFast Xpos, Ypos, Source, rRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

'This sub will blit any objects such as the player, enemies etc..with clipping(dont show what doesnt need to be shown...)
Public Sub BltObject(Destination As DirectDrawSurface7, Source As DirectDrawSurface7, Xpos As Integer, Ypos As Integer, Direction As Integer, Top, Left, Width, Height)
    
    Dim rRect As RECT
    Dim PovX As Integer
    Dim PovY As Integer
    Dim RTop As Integer
    
    RTop = Top
    If Xpos >= 320 - Width Then Width = 320 - Xpos
    If Xpos < 0 Then
        Left = Abs(Xpos)
        Width = Width - Left
        PovX = Abs(Xpos)
    End If
    If Ypos >= 320 - Height Then Height = 320 - Ypos
    If Ypos < 0 Then
        RTop = Top + Abs(Ypos)
        Height = (Height - Abs(Ypos))
        PovY = Abs(Ypos)
    End If
        
    
    rRect.Top = RTop
    rRect.Left = Left
    rRect.Bottom = RTop + Height
    rRect.Right = Left + Width
    
    Destination.BltFast Xpos + PovX, Ypos + PovY, Source, rRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub
Public Sub ShutDown()
    
    dd.RestoreDisplayMode
    dd.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
    'lose the keyboard
    diDEV.Unacquire
    
    Set surChars.Pic = Nothing
    Set surTiles.Pic = Nothing
    Set DBuffer2.Surface = Nothing
    Set DBuffer3.Surface = Nothing
    Set surMessageBox.Pic = Nothing
    Set primary = Nothing
    Set backbuffer = Nothing
    Set dd = Nothing
    Set dx = Nothing
    RemoveFont (lngDC) 'remove the font
    Performance.Stop Segment, Nothing, 0, 0
    Performance.CloseDown 'stop the music
    End
End Sub

Public Function GetBitmapInfo(FileName As String, Optional GetHeight As Boolean, Optional GetWidth As Boolean) As Integer
    'gets the bitmaps information...such as its heigh and width
    Open FileName For Random As #1 Len = Len(BITMAP)
        Get #1, 1, BITMAP
    Close #1
    
    If GetHeight = True Then
        GetBitmapInfo = CInt(BITMAP.InfoHeader.biHeight)
    Else
        GetBitmapInfo = CInt(BITMAP.InfoHeader.biWidth)
    End If
End Function

'Thanks to Jim Camel for writing a great tutorial on loading .pngs to DirectDraw surfaces!
Public Sub LoadFromFreeImage(FileName As String, Surface As DirectDrawSurface7, MakeTrans As Boolean, Optional TransColorHigh As Long, Optional TransColorLow As Long)
'We need to check what Image Format is being used, so we can tell
'if FreeImage supports it
Dim ImageFormat As FREE_IMAGE_FORMAT
Dim TransKey As DDCOLORKEY
'Some holders for DIB's (Device Independent Bitmaps)
Dim DIB         As Long
Dim DIB_New     As Long
'A holder for our loaded BITMAP
Dim DC As Long
Dim hBITMAP     As Long

'Like before we need to create a surface
'And get the Device Context from it
Dim ddsd As DDSURFACEDESC2

  
    'Clear our surface
    Set Surface = Nothing
    
    'Initialise the FreeImage library
    FreeImage_Initialise
    
    'Find out what kind of filetype we're dealing with
    ImageFormat = FreeImage_GetFileType(FileName)
    
    'We can only deal with the image if it's a known format
    If ImageFormat <> FIF_UNKNOWN Then
        
        'Load the Device Independent Bitmap
        DIB = FreeImage_Load(ImageFormat, FileName, 0)
        
        'If the loading failed, the DIB would be 0
        
        If DIB <> 0 Then
            
            'Get a BMP from the DIB
            If FreeImage_BMP_From_DIB(DIB, hBITMAP) = True Then
                
                'Create our surface. Once again we need to specify
                'the size of the surface
                ddsd.lFlags = DDSD_WIDTH Or DDSD_HEIGHT
                ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                'However, this time we can use our bitmap and ask
                'FreeImage to find the width and height of it
                ddsd.lWidth = FreeImage_GetBitmapWidth(hBITMAP)
                ddsd.lHeight = FreeImage_GetBitmapHeight(hBITMAP)
                'Creating it is the easy part
                Set Surface = dd.CreateSurface(ddsd)
                
                'Get our Device Context, as before
                DC = Surface.GetDC
                
                'The next few steps I don't fully understand, but was written
                'Following the examples given in the FreeImage module
                'First we render the bitmap into our Surfaces' DC
                FreeImage_RenderBitmap hBITMAP, DC
                'Then we convert the Bitmap back into a DIB (why I'm not sure)
                If FreeImage_DIB_From_BMP(hBITMAP, DIB_New) = True Then
                    'And then render the Device Independent Bitmap
                    FreeImage_RenderDIB DIB_New, DC
                    'And finally free the new Device Independent Bitmap
                    FreeImage_Free DIB_New
                End If
                'Release our Surface's DC and unlock it
                Surface.ReleaseDC DC
                'Delete our Bitmap
                DeleteObject hBITMAP
            End If
            'Unload our original DIB
            FreeImage_Unload DIB
        End If
    End If
    If MakeTrans = True Then
        TransKey.high = TransColorHigh
        TransKey.low = TransColorLow

        Surface.SetColorKey DDCKEY_SRCBLT, TransKey
    End If
    'And finally, deinitialise the FreeImage library. We're finished!
    FreeImage_DeInitialise
End Sub

'Checks to see if a specific key has been pressed
Public Function CheckKey(Key As Integer) As Boolean
    diDEV.GetDeviceStateKeyboard diState
    If diState.Key(Key) <> 0 Then
        CheckKey = True
    Else
        CheckKey = False
    End If
End Function







