Attribute VB_Name = "LoadPNG"


Private Sub LoadFromFreeImage(FileName As String, Surface As DirectDrawSurface7)
'We need to check what Image Format is being used, so we can tell
'if FreeImage supports it
Dim ImageFormat As FREE_IMAGE_FORMAT
'Some holders for DIB's (Device Independent Bitmaps)
Dim DIB         As Long
Dim DIB_New     As Long
'A holder for our loaded BITMAP
Dim hBITMAP     As Long

'Like before we need to create a surface
'And get the Device Context from it

Dim DC As Long
  
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
                ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
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
    
    'And finally, deinitialise the FreeImage library. We're finished!
    FreeImage_DeInitialise
End Sub

