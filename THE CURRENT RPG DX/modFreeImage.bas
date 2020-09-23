Attribute VB_Name = "modFreeImage"
Option Explicit


'=============================================================================================================
'
' modFreeImage Module
' -------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone )
'
' Created On  : March 12, 2002
' Last Update : May 28, 2003
'
' VB Versions : 5.0 / 6.0
'
' Requires    : FreeImage.dll [v2.5.2 - v2.6.1] (FreeImage library by Floris van den Berg - flvdberg@wxs.nl)
'                 - http://freeimage.sourceforge.net/
'                 - http://www.sourceforge.net/projects/freeimage/
'                 - http://groups.yahoo.com/group/FreeImageDev/
'
' Description : This module wraps the exported functions and functionality of the FreeImage library
'               making usable to Visual Basic 5/6 programmers.  This module also adds on a few extra functions
'               that make using the FreeImage.dll library easier within Visual Basic.
'
' Supported   : .BMP     [IN/OUT] Windows or OS/2 Bitmap
'               .RLE     [IN/OUT] Bitmap with run-length encoding (RLE4/RLE8 in, RLE8 out)
'               .WBMP    [IN/OUT] Wireless Bitmap - A bitmap type used by cell phones, etc.  Monochrome (B&W) Only.
'               .PBM     [IN/OUT] Portable Bitmap - Black & white ASCII based bitmap format
'               .PGM     [IN/OUT] Portable Greymap - Grayscale ASCII based bitmap format
'               .PPM     [IN/OUT] Portable Pixmap - High color ASCII based bitmap format
'               .PBMRAW  [IN/OUT] (Binary version of .PBM)
'               .PGMRAW  [IN/OUT] (Binary version of .PGM)
'               .PPMRAW  [IN/OUT] (Binary version of .PPM)
'               .ICO     [IN    ] Windows Icon
'               .PSD     [IN    ] Adobe Photoshop Document
'               .JPEG    [IN/OUT] Joint Photographic Experts Group (JFIF Compliant) - Popular high color bitmap format used on the WWW
'               .JBIG    [IN/OUT] Joint Bi-level Image experts Group.  Cousin of JPEG, but B&W (REQUIRES PLUG-IN)
'               .JNG     [IN    ] JPEG Network Graphics.  Cousin of MNG, but uses JPEG compression
'               .PNG     [IN/OUT] Portable Network Graphics - The answer to GIF when UNISYS became intrusive about their LZW compression algorithm
'               .MNG     [IN    ] Multiple Network Graphics - PNG's answer to animated GIF.
'               .PCX     [IN    ] Zsoft Paintbrush
'               .TIFF    [IN/OUT] Tag Image File Format - A widely used format for storing image data
'               .TARGA   [IN/OUT] Tagged Image File Format - Popular bitmap format on the Amiga and other computers
'               .KOALA   [IN    ] KOALA files were popular on the commodore 64. Added for nostalgic reasons
'               .PCD     [IN    ] Kodak PhotoCD. The PhotoCD format was developed by Kodak as an alternative to analog photography. Natively supported by most DVD players and CD-I.
'               .RAS     [IN    ] Sun Raster File - Popular among Solaris computers
'               .IFF     [IN    ] Interchanged File Format - Designed by Electronic Arts as an Amiga image storage format
'               .LBM     [IN    ] LBM was created for the Deluxe Paint package, and is essentially the same as IFF. (still part of FreeImage.dll for backwards compatibility)
'               .CUT     [IN    ] Dr.Halo bitmap - Pretty popular some years ago
'               .XBM     [IN    ] X11 Bitmap Format
'
' NOTE        : Within this code you'll see the following return types:
'                 - FIBITMAP      = A handle to a Windows Device Independant Bitmap (DIB)
'                 - FIMULTIBITMAP = A handle to a "Multi-Page" bitmap (meaning an image that contains multiple
'                                   pages... like TIFF's sometimes are).  Refer to "FreeImage_MultiPage_Load"
'                 - DIB           = A handle to a Windows Device Independant Bitmap (DIB)
'                 - RGBQUAD       = A single RGBQUAD structure (type) represents the RED, GREE, & BLUE colors
'                                   make up a single color.  An array of RGBQUAD's is commonly used as a palette
'
' Example Use :
'
'-------------------------------------------------------------------------------------------------------------
'
'  Const TEST_FILE As String = "C:\TEST.BMP"
'  Dim ImageFormat As FREE_IMAGE_FORMAT
'  Dim DIB         As Long
'  Dim DIB_New     As Long
'  Dim hBITMAP     As Long
'  Me.Show
'  Me.AutoRedraw = True
'  FreeImage_Initialise
'  ImageFormat = FreeImage_GetFileType(TEST_FILE, 16)
'  If ImageFormat <> FIF_UNKNOWN Then
'    DIB = FreeImage_Load(ImageFormat, TEST_FILE, 0)
'    If DIB <> 0 Then
'      If FreeImage_BMP_From_DIB(DIB, hBITMAP) = True Then
'        FreeImage_RenderBitmap hBITMAP, Me.hDC
'        If FreeImage_DIB_From_BMP(hBITMAP, DIB_New) = True Then
'          FreeImage_RenderDIB DIB_New, Me.hDC
'          FreeImage_Free DIB_New
'        End If
'        DeleteObject hBITMAP
'      End If
'      FreeImage_Unload DIB
'    End If
'  End If
'  FreeImage_DeInitialise
'
'-------------------------------------------------------------------------------------------------------------
'
'  Const INPUT_FILE     As String = "C:\TEST.BMP"
'  Const OUTPUT_FILE    As String = "C:\NEW.JPG"
'  Const OUTPUT_QUALITY As Long = 85
'  Dim DIB         As Long
'  Dim DIB2        As Long
'  Dim ImageFormat As FREE_IMAGE_FORMAT
'  Dim objPic      As StdPicture
'  FreeImage_Initialise
'  Set objPic = LoadPicture(INPUT_FILE)
'  If Not objPic Is Nothing Then
'    If FreeImage_DIB_From_BMP(objPic.Handle, DIB) = True Then
'      DIB2 = FreeImage_ConvertTo24Bits(DIB)
'      If Dir(OUTPUT_FILE, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then Kill OUTPUT_FILE
'      If FreeImage_Save(FIF_JPEG, DIB2, OUTPUT_FILE, OUTPUT_QUALITY) = TRUE_ Then
'         MsgBox "Successfully saved to JPEG!", vbOKOnly + vbInformation, " "
'      End If
'    End If
'  End If
'  If DIB <> 0 Then FreeImage_Unload DIB
'  If DIB2 <> 0 Then FreeImage_Unload DIB2
'  FreeImage_DeInitialise
'
'-------------------------------------------------------------------------------------------------------------
'
'  Const FILE_INPUT  As String = "C:\TEST.BMP"
'  Const FILE_OUTPUT As String = "C:\NEW.JPG"
'  Dim ImgType  As FREE_IMAGE_FORMAT
'  Dim DIB      As Long
'  Dim DIB_New  As Long
'  Dim DIB_8Bit As Long
'  Me.Show
'  Me.AutoRedraw = True
'  If Dir(FILE_OUTPUT, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then Kill FILE_OUTPUT
'  FreeImage_Initialise
'  ImgType = FreeImage_GetFileType(FILE_INPUT)
'  DIB = FreeImage_Load(ImgType, FILE_INPUT)
'  If DIB <> 0 Then
'    If FreeImage_GrayscaleDIB(DIB, DIB_New) = True Then
'      DIB_8Bit = FreeImage_ConvertTo8Bits(DIB_New)
'      FreeImage_RenderDIB DIB_8Bit, Me.hDC
'      Me.Refresh
'      FreeImage_Save FIF_JPEG, DIB_8Bit, FILE_OUTPUT
'      FreeImage_Unload DIB_New
'      FreeImage_Unload DIB_8Bit
'    End If
'    FreeImage_Unload DIB
'  End If
'  FreeImage_DeInitialise
'
'=============================================================================================================


'=============================================================================================================
' FreeImage 2
'
' Design and implementation by
' - Floris van den Berg (flvdberg@wxs.nl)
'
' Contributors:
' - Adam Gates (radad@xoasis.com)
' - Alex Kwak
' - Alexander Dymerets (sashad@te.net.ua)
' - Detlev Vendt (detlev.vendt@brillit.de)
' - Hervé Drolon (drolon@infonie.fr)
' - Jan L. Nauta (jln@magentammt.com)
' - Jani Kajala (janik@remedy.fi)
' - Juergen Riecker (j.riecker@gmx.de)
' - Laurent Rocher (rocherl@club-internet.fr)
' - Luca Piergentili (l.pierge@terra.es)
' - Machiel ten Brinke (brinkem@uni-one.nl)
' - Markus Loibl (markus.loibl@epost.de)
' - Martin Weber (martweb@gmx.net)
' - Matthias Wandel (mwandel@rim.net)
'
' This file is part of FreeImage 2
'
' COVERED CODE IS PROVIDED UNDER THIS LICENSE ON AN "AS IS" BASIS, WITHOUT WARRANTY
' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, WITHOUT LIMITATION, WARRANTIES
' THAT THE COVERED CODE IS FREE OF DEFECTS, MERCHANTABLE, FIT FOR A PARTICULAR PURPOSE
' OR NON-INFRINGING. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE COVERED
' CODE IS WITH YOU. SHOULD ANY COVERED CODE PROVE DEFECTIVE IN ANY RESPECT, YOU (NOT
' THE INITIAL DEVELOPER OR ANY OTHER CONTRIBUTOR) ASSUME THE COST OF ANY NECESSARY
' SERVICING, REPAIR OR CORRECTION. THIS DISCLAIMER OF WARRANTY CONSTITUTES AN ESSENTIAL
' PART OF THIS LICENSE. NO USE OF ANY COVERED CODE IS AUTHORIZED HEREUNDER EXCEPT UNDER
' THIS DISCLAIMER.
'
' Use at your own risk!
'=============================================================================================================


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                       TYPE / ENUMERATION  DECLARATIONS                  XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'_____________________________________________________________________________________________________________
' Bitmap types
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' Describes a color consisting of relative intensities of red, green, and blue.
Public Type RGBQUAD
  rgbBlue     As Byte '//BYTE  Specifies the intensity of blue in the color.
  rgbGreen    As Byte '//BYTE  Specifies the intensity of green in the color.
  rgbRed      As Byte '//BYTE  Specifies the intensity of red in the color.
  rgbReserved As Byte '//BYTE  Reserved; must be zero.
End Type

' Contains information about the dimensions and color format of a device-independent bitmap (DIB)
Public Type BITMAPINFOHEADER
  biSize          As Long    '//DWORD  Specifies the number of bytes required by the structure.
  biWidth         As Long    '//LONG   Specifies the width of the bitmap, in pixels.
  biHeight        As Long    '//LONG   Specifies the height of the bitmap, in pixels. If biHeight is positive, the bitmap is a bottom-up DIB and its origin is the lower left corner. If biHeight is negative, the bitmap is a top-down DIB and its origin is the upper left corner.
  biPlanes        As Integer '//WORD   Specifies the number of planes for the target device. This value must be set to 1.
  biBitCount      As Integer '//WORD   Specifies the number of bits per pixel. This value must be 1, 4, 8, 16, 24, or 32.
  biCompression   As Long    '//DWORD  Specifies the type of compression for a compressed bottom-up bitmap (top-down DIBs cannot be compressed). It can be one of the following values: BI_RGB, BI_RLE8, BI_RLE4, BI_BITFIELDS
  biSizeImage     As Long    '//DWORD  Specifies the size, in bytes, of the image. This may be set to 0 for BI_RGB bitmaps.
  biXPelsPerMeter As Long    '//LONG   Specifies the horizontal resolution, in pixels per meter, of the target device for the bitmap. An application can use this value to select a bitmap from a resource group that best matches the characteristics of the current device.
  biYPelsPerMeter As Long    '//LONG   Specifies the vertical resolution, in pixels per meter, of the target device for the bitmap.
  biClrUsed       As Long    '//DWORD  Specifies the number of color indices in the color table that are actually used by the bitmap. If this value is zero, the bitmap uses the maximum number of colors corresponding to the value of the biBitCount member for the compression mode specified by biCompression.
                             '         If biClrUsed is nonzero and the biBitCount member is less than 16, the biClrUsed member specifies the actual number of colors the graphics engine or device driver accesses. If biBitCount is 16 or greater, then biClrUsed member specifies the size of the color table used to optimize performance of Windows color palettes. If biBitCount equals 16 or 32, the optimal color palette starts immediately following the three doubleword masks.
                             '         If the bitmap is a packed bitmap (a bitmap in which the bitmap array immediately follows the BITMAPINFO header and which is referenced by a single pointer), the biClrUsed member must be either 0 or the actual size of the color table.
  biClrImportant  As Long    '//DWORD  Specifies the number of color indices that are considered important for displaying the bitmap. If this value is zero, all colors are important.
End Type

' Defines the dimensions and color information for a Windows device-independent bitmap (DIB)
Public Type BITMAPINFO
  bmiHeader   As BITMAPINFOHEADER '// BITMAPINFOHEADER  Specifies a BITMAPINFOHEADER structure that contains information about the dimensions and color format of a DIB.
  bmiColors() As RGBQUAD          '// RGBQUAD           Specifies an array of RGBQUAD or doubleword data types that define the colors in the bitmap.
End Type

' Defines the type, width, height, color format, and bit values of a bitmap.
Public Type BITMAP
  bmType       As Long    '// LONG   Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    '// LONG   Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    '// LONG   Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    '// LONG   Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer '// WORD   Specifies the count of color planes.
  bmBitsPixel  As Integer '// WORD   Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    '// LPVOID Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

' FreeImageIO is a structure containing pointers to 4 functions: a read function, write function,
' seek function and tell function.  All these functions have to be implemented so that data is
' delivered.  The handle representing the data is made abstract as well and is named fi_handle.
Public Type FreeImageIO
  read_proc  As Long '// FI_ReadProc   // pointer to the function used to read data
  write_proc As Long '// FI_WriteProc  // pointer to the function used to write data
  seek_proc  As Long '// FI_SeekProc   // pointer to the function used to seek
  tell_proc  As Long '// FI_TellProc   // pointer to the function used to aquire the current position
End Type

' ICC profile support
Public Type FIICCPROFILE
   flags  As Integer '// WORD  - info flag
   Size   As Long    '// DWORD - profile's size measured in bytes
   Data() As Byte    '// void  - points to a block of contiguous memory containing the profile
End Type

'_____________________________________________________________________________________________________________
' Other Types
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯


Public Type Plugin
  Format_Proc                As Long '// FI_FormatProc
  Description_Proc           As Long '// FI_DescriptionProc
  Extension_Proc             As Long '// FI_ExtensionListProc
  RegExpr_Proc               As Long '// FI_RegExprProc
  Open_Proc                  As Long '// FI_OpenProc
  Close_Proc                 As Long '// FI_CloseProc
  PageCount_Proc             As Long '// FI_PageCountProc
  PageCapability_Proc        As Long '// FI_PageCapabilityProc
  Load_Proc                  As Long '// FI_LoadProc
  Save_Proc                  As Long '// FI_SaveProc
  Validate_Proc              As Long '// FI_ValidateProc
  Mime_Proc                  As Long '// FI_MimeProc
  Supports_Export_BPP_Proc   As Long '// FI_SupportsExportBPPProc
  Supports_ICC_Profiles_Proc As Long '// FI_SupportsICCProfilesProc
End Type

'_____________________________________________________________________________________________________________
' Important enums
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Public Enum BOOL
  TRUE_ = 1  '// TRUE  in C/C++ = 1 ... TRUE in Visual Basic = -1
  FALSE_ = 0 '// FALSE in C/C++ = 0 ... FALSE in Visual Basic = 0
End Enum

Public Enum FREE_IMAGE_FORMAT
   FIF_UNKNOWN = -1  '// Unknown format
   FIF_BMP = 0       '// Windows or OS/2 Bitmap File (*.BMP)
   FIF_ICO           '// Windows Icon (*.ICO)
   FIF_JPEG          '// Independent JPEG Group (*.JPG)
   FIF_JNG           '// JPEG Network Graphics (*.JNG)
   FIF_KOALA         '// Commodore 64 Koala format (*.KOA)
   FIF_LBM           '// Amiga IFF (*.IFF, *.LBM)
   FIF_MNG           '// Multiple Network Graphics (*.MNG)
   FIF_PBM           '// Portable Bitmap (ASCII) (*.PBM)
   FIF_PBMRAW        '// Portable Bitmap (BINARY) (*.PBM / *.PBMRAW)
   FIF_PCD           '// Kodak PhotoCD (*.PCD)
   FIF_PCX           '// PCX bitmap format (*.PCX)
   FIF_PGM           '// Portable Graymap (ASCII) (*.PGM)
   FIF_PGMRAW        '// Portable Graymap (BINARY) (*.PGM / *.PGMRAW)
   FIF_PNG           '// Portable Network Graphics (*.PNG)
   FIF_PPM           '// Portable Pixelmap (ASCII) (*.PPM)
   FIF_PPMRAW        '// Portable Pixelmap (BINARY) (*.PPM / *.PPMRAW)
   FIF_RAS           '// Sun Rasterfile (*.RAS)
   FIF_TARGA         '// Targa files (*.TGA)
   FIF_TIFF          '// Tagged Image File Format (*.TIFF)
   FIF_WBMP          '// Wireless Bitmap (*.WBMP)
   FIF_PSD           '// Adobe Photoshop Document (*.PSD)
   FIF_CUT           '// Dr. Halo (*.CUT)
   FIF_IFF = FIF_LBM '// Amiga IFF (*.IFF, *.LBM)
   FIF_XBM           '// X11 Bitmap Format (*.XBM)
End Enum

Public Enum FREE_IMAGE_COLOR_TYPE
   FIC_MINISWHITE = 0 '// Monochrome bitmap: first palette entry is black (1 bit)
                      '   Palettized bitmap: grayscale palette (8 bit)
   FIC_MINISBLACK = 1 '// Monochrome bitmap: first palette entry is white (1 bit)
   FIC_RGB = 2        '// Palettized bitmap (1, 4 or 8 bit)
   FIC_PALETTE = 3    '// High-color bitmap (16, 24 or 32 bit)
   FIC_RGBALPHA = 4   '// High-color bitmap with an alpha channel (32 bit only)
   FIC_CMYK = 5       '// CMYK bitmap (32 bit only)
End Enum

Public Enum FREE_IMAGE_QUANTIZE
  FIQ_WUQUANT = 0 '// Xiaolin Wu color quantization algorithm
  FIQ_NNQUANT = 1 '// NeuQuant neural-net quantization algorithm by Anthony Dekker
End Enum

Public Enum FREE_IMAGE_DITHER
   FID_FS = 0           '// Floyd & Steinberg error diffusion algorithm
   FID_BAYER4x4 = 1     '// Bayer ordered dispersed dot dithering (order 2 [4x4] dithering matrix)
   FID_BAYER8x8 = 2     '// Bayer ordered dispersed dot dithering (order 3 [8x8] dithering matrix)
   FID_CLUSTER6x6 = 3   '// Ordered clustered dot dithering (order 3 [6x6] matrix)
   FID_CLUSTER8x8 = 4   '// Ordered clustered dot dithering (order 4 [8x8] matrix)
   FID_CLUSTER16x16 = 5 '// Ordered clustered dot dithering (order 8 [16x16] matrix)
End Enum

Public Type FreeImage
   Allocate_Proc                      As Long
   Unload_Proc                        As Long
   Free_Proc                          As Long
   Get_Colors_Used_Proc               As Long
   Get_Bits_Proc                      As Long
   Get_Bits_row_col_Proc              As Long
   Get_Scanline_Proc                  As Long
   Get_BPP_Proc                       As Long
   Get_Width_Proc                     As Long
   Get_Height_Proc                    As Long
   Get_Line_Proc                      As Long
   Get_Pitch_Proc                     As Long
   Get_DIB_Size_Proc                  As Long
   Get_Palette_Proc                   As Long
   Get_Dots_Per_Meter_X_Proc          As Long
   Get_Dots_Per_Meter_Y_Proc          As Long
   Get_Info_Header_Proc               As Long
   Get_Info_Proc                      As Long
   get_icc_profile_proc               As Long
   create_icc_profile_proc            As Long
   destroy_icc_profile_proc           As Long
   Get_Color_Type_Proc                As Long
   Get_Red_Mask_Proc                  As Long
   Get_Green_Mask_Proc                As Long
   Get_Blue_Mask_Proc                 As Long
   Get_Transparency_Count_Proc        As Long
   Get_Transparency_Table_Proc        As Long
   Set_Transparency_Table_Proc        As Long
   Is_Transparent_Proc                As Long
   Set_Transparent_Proc               As Long
   Output_Message_Proc                As Long
   Convert_Line1to8_Proc              As Long
   Convert_Line_4to8_Proc             As Long
   Convert_Line_16to8_555_Proc        As Long
   Convert_Line_16to8_565_Proc        As Long
   Convert_Line_24to8_Proc            As Long
   Convert_Line_32to8_Proc            As Long
   Convert_Line_1to16_555_Proc        As Long
   Convert_Line_4to16_555_Proc        As Long
   Convert_Line_8to16_555_Proc        As Long
   Convert_Line_16_565_to_16_555_Proc As Long
   Convert_Line_24to16_555_Proc       As Long
   Convert_Line_32to16_555_Proc       As Long
   Convert_Line_1to16_565_Proc        As Long
   Convert_Line_4to16_565_Proc        As Long
   Convert_Line_8to16_565_Proc        As Long
   Convert_Line_16_555_to_16_565_Proc As Long
   Convert_Line_24to16_565_Proc       As Long
   Convert_Line_32to16_565_Proc       As Long
   Convert_Line_1to24_Proc            As Long
   Convert_Line_4to24_Proc            As Long
   Convert_Line_8to24_Proc            As Long
   Convert_Line_16to24_555_Proc       As Long
   Convert_Line_16to24_565_Proc       As Long
   Convert_Line_32to24_Proc           As Long
   Convert_Line_1to32_Proc            As Long
   Convert_Line_4to32_Proc            As Long
   Convert_Line_8to32_Proc            As Long
   Convert_Line_16to32_555_Proc       As Long
   Convert_Line_16to32_565_Proc       As Long
   Convert_Line_24to32_Proc           As Long
End Type


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                          CONSTANTS  DECLARATIONS                        XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'_____________________________________________________________________________________________________________
' Load/Save flag constants
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' NOTE: If you are saving a JPEG image, you can pass the quality of the JPEG image (1 to 100) as the "flags" parameter
Public Const BMP_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const BMP_SAVE_RLE        As Long = 1     ' [  /Out] Saves the BITMAP with RLE8 run-length encoding
Public Const CUT_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const ICO_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const ICO_FIRST           As Long = 0     ' [In    ] Loads the first bitmap in the icon
Public Const ICO_SECOND          As Long = 0     ' [In    ] Loads the second bitmap in the icon
Public Const ICO_THIRD           As Long = 0     ' [In    ] Loads the third bitmap in the icon
Public Const IFF_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const JPEG_DEFAULT        As Long = 0     ' [In/Out] Default flag
Public Const JPEG_FAST           As Long = 1     ' [In    ] Loads the file as fast as possible, sacrificing some quality
Public Const JPEG_ACCURATE       As Long = 2     ' [In    ] Loads the file with the best quality, sacrificing some speed
Public Const JPEG_QUALITYSUPERB  As Long = &H80  ' [   Out] Saves with superb quality
Public Const JPEG_QUALITYGOOD    As Long = &H100 ' [   Out] Saves with good quality
Public Const JPEG_QUALITYNORMAL  As Long = &H200 ' [   Out] Saves with normal quality
Public Const JPEG_QUALITYAVERAGE As Long = &H400 ' [   Out] Saves with average quality
Public Const JPEG_QUALITYBAD     As Long = &H800 ' [   Out] Saves with bad quality
Public Const KOALA_DEFAULT       As Long = 0     ' [In/Out] Default flag
Public Const LBM_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const MNG_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const PCD_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const PCD_BASE            As Long = 1     ' [In    ] A PhotoCD picture comes in many sizes. This flag will load the one sized 768 x 512
Public Const PCD_BASEDIV4        As Long = 2     ' [In    ] This flag will load the bitmap sized 384 x 256
Public Const PCD_BASEDIV16       As Long = 3     ' [In    ] This flag will load the bitmap sized 192 x 128
Public Const PCX_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const PNG_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const PNG_IGNOREGAMMA     As Long = 1     ' [In    ] Avoid gamma correction
Public Const PNM_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const PNM_SAVE_RAW        As Long = 0     ' [   Out] If set the writer saves in RAW format (i.e. P4, P5 or P6)
Public Const PNM_SAVE_ASCII      As Long = 1     ' [   Out] If set the writer saves in ASCII format (i.e. P1, P2 or P3)
Public Const RAS_DEFAULT         As Long = 0     ' [In/Out] Default flag
Public Const TARGA_DEFAULT       As Long = 0     ' [In/Out] Default flag
Public Const TARGA_LOAD_RGB888   As Long = 1     ' [In    ] If set the loader converts RGB555 and ARGB8888 -> RGB888.
'Public Const TARGA_LOAD_RGB555  As Long = 2     ' [******] This flag is obsolete!
Public Const TIFF_DEFAULT        As Long = 0     ' [In/Out] Default flag
Public Const TIFF_CMYK           As Long = &H1   ' [In/Out] Reads/stores tags for separated CMYK (use | to combine with compression flags)
Public Const TIFF_PACKBITS       As Long = &H100 ' [  /Out] Save using PACKBITS compression
Public Const TIFF_DEFLATE        As Long = &H200 ' [  /Out] Save using DEFLATE compression
Public Const TIFF_ADOBE_DEFLATE  As Long = &H400 ' [  /Out] Save using ADOBE DEFLATE compression
Public Const TIFF_NONE           As Long = &H800 ' [  /Out] Save without any compression
Public Const WBMP_DEFAULT        As Long = 0     ' [In/Out] Default flag
Public Const PSD_DEFAULT         As Long = 0     ' [In/Out] Default flag

'_____________________________________________________________________________________________________________
' Constants - ICC profile support
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Const FIICC_DEFAULT       As Long = &H0
Public Const FIICC_COLOR_IS_CMYK As Long = &H1

'_____________________________________________________________________________________________________________
' Other Constants
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Const MAX_PATH            As Long = 260
Public Const BI_RGB              As Long = 0 ' An uncompressed format.
Public Const BI_RLE8             As Long = 1 ' A run-length encoded (RLE) format for bitmaps with 8 bits per pixel. The compression format is a two-byte format consisting of a count byte followed by a byte containing a color index. For more information, see the following Remarks section.
Public Const BI_RLE4             As Long = 2 ' An RLE format for bitmaps with 4 bits per pixel. The compression format is a two-byte format consisting of a count byte followed by two word-length color indices. For more information, see the following Remarks section.
Public Const BI_BITFIELDS        As Long = 3 ' Specifies that the bitmap is not compressed and that the color table consists of three doubleword color masks that specify the red, green, and blue components, respectively, of each pixel. This is valid when used with 16- and 32-bits-per-pixel bitmaps.

'_____________________________________________________________________________________________________________
' Local Variables
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private LastError_ErrDesc        As String


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                        SUB / FUNCTION  DECLARATIONS                     XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Init/Error routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Initialises the library.  Calling the "FreeImage_Initialise" function below will enable error trapping, so
' I've changed the names so that will be used instead.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_Initialise(BOOL load_local_plugins_only FI_DEFAULT(FALSE));
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_Init Lib "FREEIMAGE.DLL" Alias "_FreeImage_Initialise@4" (Optional ByVal load_local_plugins_only As BOOL = FALSE_)

'-------------------------------------------------------------------------------------------------------------
' DeInitialises the library.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_DeInitialise();
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_DeInitialise Lib "FREEIMAGE.DLL" Alias "_FreeImage_DeInitialise@0" ()

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Version routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Returns a POINTER to a string that contains the version string of the library.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char *DLL_CALLCONV FreeImage_GetVersion();
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetVer Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetVersion@0" () As Long

'-------------------------------------------------------------------------------------------------------------
' Returns a POINTER to a string containing a short copyright message you can include in your program.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char *DLL_CALLCONV FreeImage_GetCopyrightMessage();
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetCopyrightMsg Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetCopyrightMessage@0" () As Long

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Message output functions
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Callback function definition
'-------------------------------------------------------------------------------------------------------------
'typedef void (*FreeImage_OutputMessageFunction)(FREE_IMAGE_FORMAT fif, const char *msg);
'-------------------------------------------------------------------------------------------------------------
' (See the public function "FI_OutputMessageProc")

'-------------------------------------------------------------------------------------------------------------
' Sets the Message Proc equal to the function you specify (in Visual Basic, use the "AddressOf" operator to
' get the memory address of a PUBLIC FUNCTION [which must be located in a standard VB module] to act as this
' proc)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_SetOutputMessage(FreeImage_OutputMessageFunction omf);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_SetOutputMessage Lib "FREEIMAGE.DLL" Alias "_FreeImage_SetOutputMessage@4" (ByVal Msg_Callback_Function As Long)

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Allocate/Unload routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Allocates a new FreeImage bitmap using the given width, height and bits per pixel and optional red, green
' and blue mask. The mask is stored in BGR order.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_Allocate(int width, int height, int bpp, unsigned red_mask FI_DEFAULT(0), unsigned green_mask FI_DEFAULT(0), unsigned blue_mask FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Allocate Lib "FREEIMAGE.DLL" Alias "_FreeImage_Allocate@24" (ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long, Optional ByVal Red_Mask As Long = 0, Optional ByVal Green_Mask As Long = 0, Optional ByVal Blue_Mask As Long = 0) As Long

'-------------------------------------------------------------------------------------------------------------
' Disposes the given bitmap from memory.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_Free(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_Free Lib "FREEIMAGE.DLL" Alias "_FreeImage_Free@4" (ByVal DIB As Long)

'-------------------------------------------------------------------------------------------------------------
' Alias for FreeImage_Free.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_Unload(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_Unload Lib "FREEIMAGE.DLL" Alias "_FreeImage_Unload@4" (ByVal DIB As Long)

'-------------------------------------------------------------------------------------------------------------
' Makes an exact copy of an existing bitmap in memory
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP * DLL_CALLCONV FreeImage_Clone(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Clone Lib "FREEIMAGE.DLL" Alias "_FreeImage_Clone@4" (ByVal DIB As Long) As Long

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Plugin Interface
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Registers a plugin which is residing in the programmer's own sourcebase (not a DLL)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_RegisterLocalPlugin(FI_InitProc proc_address, const char *format FI_DEFAULT(0), const char *description FI_DEFAULT(0), const char *extension FI_DEFAULT(0), const char *regexpr FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_RegisterLocalPlugin Lib "FREEIMAGE.DLL" Alias "_FreeImage_RegisterLocalPlugin@20" (ByVal Proc_Address As Long, Optional ByVal Format As String = vbNullString, Optional ByVal Description As String = vbNullString, Optional ByVal Extention As String = vbNullString, Optional ByVal RegExpr As String = vbNullString) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Register a plugin residing in a DLL
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_RegisterExternalPlugin(const char *path, const char *format FI_DEFAULT(0), const char *description FI_DEFAULT(0), const char *extension FI_DEFAULT(0), const char *regexpr FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_RegisterExternalPlugin Lib "FREEIMAGE.DLL" Alias "_FreeImage_RegisterExternalPlugin@20" (ByVal Path As String, Optional ByVal Format As String = vbNullString, Optional ByVal Description As String = vbNullString, Optional ByVal Extention As String = vbNullString, Optional ByVal RegExpr As String = vbNullString) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Returns the number of supported bitmap formats.
'-------------------------------------------------------------------------------------------------------------
'DLL_API int DLL_CALLCONV FreeImage_GetFIFCount();
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFCount Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFCount@0" () As Long

'-------------------------------------------------------------------------------------------------------------
' Disabled and enables a plugin. If you disable a plug-in (file format), is  will no longer be possible
' to load and save using that plug-in.
'-------------------------------------------------------------------------------------------------------------
'DLL_API int DLL_CALLCONV FreeImage_SetPluginEnabled(FREE_IMAGE_FORMAT fif, BOOL enable);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SetPluginEnabled Lib "FREEIMAGE.DLL" Alias "_FreeImage_SetPluginEnabled@8" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal Enabled As BOOL) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the status of the plugin (Enabled / Disabled)
'-------------------------------------------------------------------------------------------------------------
'DLL_API int DLL_CALLCONV FreeImage_IsPluginEnabled(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_IsPluginEnabled Lib "FREEIMAGE.DLL" Alias "_FreeImage_IsPluginEnabled@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the FreeImage format ID for the given format string.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFIFFromFormat(const char *format);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFFromFormat Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFFromFormat@4" (ByVal Format As String) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Returns a plugin constant (FIF_BMP, FIF_TIFF, etc.) from a mime string (file extention)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFIFFromMime(const char *mime);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFFromMime Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFFromMime@4" (ByVal Mime As String) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Returns a format string for the given FreeImage Format ID.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char *DLL_CALLCONV FreeImage_GetFormatFromFIF(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFormatFromFIF Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFormatFromFIF@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As String

'-------------------------------------------------------------------------------------------------------------
' Returns a list of possible file extension for the given FreeImage Format ID. The extension list is a string
' containing one or more comma separated file extensions.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char *DLL_CALLCONV FreeImage_GetFIFExtensionList(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFExtensionList Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFExtensionList@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As String

'-------------------------------------------------------------------------------------------------------------
' Returns a descriptive string for the given FreeImage Format ID.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char *DLL_CALLCONV FreeImage_GetFIFDescription(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFDescription Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFDescription@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As String

'-------------------------------------------------------------------------------------------------------------
' Returns a regular expression for the given FreeImage Format ID, assisting external libraries to identify the
' bitmap format. Currently FreeImageQt uses this function.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char * DLL_CALLCONV FreeImage_GetFIFRegExpr(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFRegExpr Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFDescription@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As String

'-------------------------------------------------------------------------------------------------------------
' Tries to identify a bitmap type by looking at the filename or a file extension. FreeImage_GetFIFFromFilename
' returns a valid FreeImage Format ID on success and FIF_UNKNOWN on failure.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFIFFromFilename(const char *filename);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFIFFromFilename Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFIFFromFilename@4" (ByVal FileName As String) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Returns true if the plugin belonging to this FreeImage Format ID supports reading, false otherwise.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_FIFSupportsReading(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_FIFSupportsReading Lib "FREEIMAGE.DLL" Alias "_FreeImage_FIFSupportsReading@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Returns true if the plugin belonging to this FreeImage Format ID supports writing, false otherwise.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_FIFSupportsWriting(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_FIFSupportsWriting Lib "FREEIMAGE.DLL" Alias "_FreeImage_FIFSupportsWriting@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Returns TRUE if the plug-in can export a bitmap in the desired color depth (bits per pixel)
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_FIFSupportsExportBPP(FREE_IMAGE_FORMAT fif, int bpp);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_FIFSupportsExportBPP Lib "FREEIMAGE.DLL" Alias "_FreeImage_FIFSupportsExportBPP@8" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal BitsPerPixel As Long) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Returns TRUE if the plugin belonging to the given FREE_IMAGE_FORMAT can load or save an ICC profile,
' FALSE otherwise.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_FIFSupportsICCProfiles(FREE_IMAGE_FORMAT fif);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_FIFSupportsICCProfiles Lib "FREEIMAGE.DLL" Alias "_FreeImage_FIFSupportsICCProfiles@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Loads the bitmap file into a FreeImage bitmap. If the bitmap is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the bitmap couldn 't be loaded, FreeImage_Load returns NULL.
'
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_Load(FREE_IMAGE_FORMAT fif, const char *filename, int flags FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Load Lib "FREEIMAGE.DLL" Alias "_FreeImage_Load@12" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal flags As Long = 0) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads a bitmap file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If the
' file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadFromHandle returns NULL.
'
' The fif parameter specifies the desired bitmap format to be loaded. The parameter accepts any valid
' FreeImage Format ID, such as FIF_BMP or FIF_JPEG, or an integer value obtained from any of the plugin
' interface functions.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadFromHandle(FREE_IMAGE_FORMAT fif, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadBMPFromHandle@12" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByRef IO_Callback_Functions As FreeImageIO, Optional ByVal flags As Long = 0) As Long

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a bitmap file of the given type.
' NOTE: If you are saving a JPEG image, you can pass the quality of the JPEG image (1 to 100) as the "flags" parameter
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_Save(FREE_IMAGE_FORMAT fif, FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Save Lib "FREEIMAGE.DLL" Alias "_FreeImage_Save@16" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = 0) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a bitmap file of the given type.
' NOTE: If you are saving a JPEG image, you can pass the quality of the JPEG image (1 to 100) as the "flags" parameter
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveToHandle(FREE_IMAGE_FORMAT fif, FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(0));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveToHandle@20" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = 0) As BOOL
                                                     
'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Multipaging interface
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' When using the multi-paging system, things are a bit different what you are used to. The bitmaps you retrieve
' from a multi-bitmap (e.g. a TIFF with multiple pages) are 'managed', which means that they are loaded and
' unloaded by the multi-page system as it sees fit. Because of this you may never call
' FreeImage_Unload/FreeImage_Free on a bitmap retrieved using FreeImage_LockPage, but instead must call
' FreeImage_unlockPage.
'
' The parameters in FreeImage_OpenMultiBitmap mean:
' 1) Selector for the bitmap plugin. Currently the only one supported is FIF_TIFF
' 2) Name of the file to be opened
' 3) True to create a new bitmap, false to open an existing one
' 4) True to open a bitmap read-only, false to allow writing. Ignored when the previous parameter is true
' 5) True if the multi-page system should keep all gathered bitmap data in memory, false to flush lazy to disc
'
' NOTE: "Sergej Kuznetsov" (sk@ghp.de) discovered that because of the way that Multipage Bitmaps work, and
' how VB creates temporary ANSI strings when passing string parameters to API's, you need to convert the
' "FileName" parameter string to ANSI from UNICODE (VB's native string format) like this:
'   strTempPath = StrConv(strFilePath, vbFromUnicode)
' then you need to pass a pointer to the string using the "StrPtr" hidden function like this:
'   DIB_Multi = FreeImage_OpenMultiBitmap(FIF_TIFF, StrPtr(strTempPath), FALSE_, FALSE_, FALSE_)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIMULTIBITMAP * DLL_CALLCONV FreeImage_OpenMultiBitmap(FREE_IMAGE_FORMAT fif, const char *filename, BOOL create_new, BOOL read_only, BOOL keep_cache_in_memory FI_DEFAULT(FALSE));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_OpenMultiBitmap Lib "FREEIMAGE.DLL" Alias "_FreeImage_OpenMultiBitmap@20" (ByVal ImageFormat As FREE_IMAGE_FORMAT, ByVal FileName As Long, ByVal CreateNew As BOOL, ByVal ReadOnly As BOOL, Optional ByVal KeepCacheInMemory As BOOL = FALSE_) As Long

'-------------------------------------------------------------------------------------------------------------
' Closes a MultiBitmap opened with a successfull call to the "FreeImage_OpenMultiBitmap" function
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_CloseMultiBitmap(FIMULTIBITMAP *bitmap);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_CloseMultiBitmap Lib "FREEIMAGE.DLL" Alias "_FreeImage_CloseMultiBitmap@4" (ByVal MultipageBitmap As Long) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Gets the total number of pages contained within a MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API int DLL_CALLCONV FreeImage_GetPageCount(FIMULTIBITMAP *bitmap);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetPageCount Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetPageCount@4" (ByVal MultipageBitmap As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Appends a page to an opened MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_AppendPage(FIMULTIBITMAP *bitmap, FIBITMAP *data);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_AppendPage Lib "FREEIMAGE.DLL" Alias "_FreeImage_AppendPage@8" (ByVal MultipageBitmap As Long, ByVal Data As Long)

'-------------------------------------------------------------------------------------------------------------
' Inserts a page into a MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_InsertPage(FIMULTIBITMAP *bitmap, int page, FIBITMAP *data);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_InsertPage Lib "FREEIMAGE.DLL" Alias "_FreeImage_InsertPage@12" (ByVal MultipageBitmap As Long, ByVal PageNumber As Long, ByVal Data As Long)

'-------------------------------------------------------------------------------------------------------------
' Deletes a page from a MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_DeletePage(FIMULTIBITMAP *bitmap, int page);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_DeletePage Lib "FREEIMAGE.DLL" Alias "_FreeImage_DeletePage@8" (ByVal MultipageBitmap As Long, ByVal PageNumber As Long)

'-------------------------------------------------------------------------------------------------------------
' Locks a specified page from a MultiBitmap so it can be saved or manipulated
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP * DLL_CALLCONV FreeImage_LockPage(FIMULTIBITMAP *bitmap, int page);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LockPage Lib "FREEIMAGE.DLL" Alias "_FreeImage_LockPage@8" (ByVal MultipageBitmap As Long, ByVal PageNumber As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Unlocks a specified page from a MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_UnlockPage(FIMULTIBITMAP *bitmap, FIBITMAP *page, BOOL changed);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_UnlockPage Lib "FREEIMAGE.DLL" Alias "_FreeImage_UnlockPage@12" (ByVal MultipageBitmap As Long, ByVal Page As Long, ByVal Changed As BOOL)

'-------------------------------------------------------------------------------------------------------------
' Relocates the specified page in a MultiBitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_MovePage(FIMULTIBITMAP *bitmap, int target, int source);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_MovePage Lib "FREEIMAGE.DLL" Alias "_FreeImage_MovePage@12" (ByVal MultipageBitmap As Long, ByVal Target As Long, ByVal Source As Long) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Returns an array of page-numbers that are currently locked in memory.  When the pages parameter is NULL,
' in count the size of the array is returned.  You can then allocate the array of the desired size and call
' FreeImage_GetLockedPageNumbers again to populate the array.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_GetLockedPageNumbers(FIMULTIBITMAP *bitmap, int *pages, int *count);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetLockedPageNumbers Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetLockedPageNumbers@12" (ByVal MultipageBitmap As Long, ByRef Pages() As Long, ByRef Count As Long) As BOOL

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Old style bitmap load routines (deprecated)
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Loads the given Windows Bitmap file or OS/2 Bitmap file into a FreeImage bitmap. If the bitmap is loaded
' successfully, memory for it is allocated and a bitmap pointer is returned. If the bitmap couldn't be loaded,
' FreeImage_LoadBMP returns NULL.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadBMP(const char *filename, int flags FI_DEFAULT(BMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadBMP Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadBMP@8" (ByVal FileName As String, Optional ByVal flags As Long = BMP_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads a Dr.Halo file (deprecated)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadCUT(const char *filename, int flags FI_DEFAULT(CUT_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadCUT Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadCUT@8" (ByVal FileName As String, Optional ByVal flags As Long = CUT_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Windows ICON file into a FreeImage bitmap. If the icon is loaded successfully, memory for
' it is allocated and a bitmap pointer is returned. If the bitmap couldn't be loaded, FreeImage_LoadICO
' returns NULL.
'
' Flags defaults to ICO_DEFAULT. The following parameters can be passed to change the behaviour of the load
' routine:
'
' FIBITMAP *dib = FreeImage_LoadICO("test.ico");
' FIBITMAP *dib = FreeImage_LoadICO("test.ico", ICO_SECOND);
'
' Flag Value:               Meaning:
' ICO_FIRST            Loads the first icon in the cabinet
' ICO_SECOND           Loads the second icon in the cabinet
' ICO_THIRD            Loads the third icon in the cabinet
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadICO(const char *filename, int flags FI_DEFAULT(ICO_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadICO Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadICO@8" (ByVal FileName As String, Optional ByVal flags As Long = ICO_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads a Interchanged File Format file (deprecated)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadIFF(const char *filename, int flags FI_DEFAULT(IFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadIFF Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadIFF@8" (ByVal FileName As String, Optional ByVal flags As Long = IFF_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given JPEG or JIF file into a FreeImage bitmap. If the file is loaded successfully, memory for it
' is allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadJPEG returns
' NULL.

' Flags defaults to JPEG_DEFAULT. The following parameters can be passed to the flags parameter to change
' the behaviour of the load routine:
'
' FIBITMAP *dib = FreeImage_LoadJPEG("test.jpg");
' FIBITMAP *dib = FreeImage_LoadJPEG("test.jpg", JPEG_ACCURATE);
'
' Flag Value:               Meaning:
' JPEG_FAST            Loads the file as fast as possible, sacrificing some quality.
' JPEG_ACCURATE        Loads the file as accurate as possible, sacrificing some speed.
' JPEG_FAST            (undocumented)
' JPEG_ACCURATE        (undocumented)
' JPEG_QUALITYSUPERB   (undocumented)
' JPEG_QUALITYGOOD     (undocumented)
' JPEG_QUALITYNORMAL   (undocumented)
' JPEG_QUALITYAVERAGE  (undocumented)
' JPEG_QUALITYBAD      (undocumented)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadJPEG(const char *filename, int flags FI_DEFAULT(JPEG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadJPEG Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadJPEG@8" (ByVal FileName As String, Optional ByVal flags As Long = JPEG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given KOALA file into a FreeImage bitmap. If the file is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the file couldn 't be loaded, FreeImage_LoadKOALA returns NULL
'
' KOALA is a 4-bit bitmap format developed for and supported by ancient commodore 64 computers and emulators.
' The format is supported for nostalgic reasons of the author.
'
' FIBITMAP *dib = FreeImage_LoadKOALA("test.koa");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadKOALA(const char *filename, int flags FI_DEFAULT(KOALA_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadKOALA Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadKOALA@8" (ByVal FileName As String, Optional ByVal flags As Long = KOALA_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Deluxe Paint LBM file into a FreeImage bitmap. If the file is loaded successfully, memory
' for it is allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadLBM
' returns NULL
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadLBM(const char *filename, int flags FI_DEFAULT(LBM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadLBM Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadLBM@8" (ByVal FileName As String, Optional ByVal flags As Long = LBM_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Multi Network Graphics (MNG) or JPEG Network Graphics (JNG) file into a FreeImage bitmap.
' If the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadMNG returns NULL
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadMNG(const char *filename, int flags FI_DEFAULT(MNG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadMNG Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadMNG@8" (ByVal FileName As String, Optional ByVal flags As Long = MNG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Kodak PhotoCD file into a FreeImage bitmap. If the file is loaded successfully, memory for
' it is allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadPCD returns
' NULL.
'
' PhotoCD images are so called cabinet files, where several versions of the same bitmap exist in one file.
' By default FreeImage loads the image marked in the flags as PCD_BASE. By using the flags parameter other
' image versions can be obtained:
'
' FIBITMAP *dib = FreeImage_LoadPCD("test.pcd");
' FIBITMAP *dib = FreeImage_LoadPCX("test.pcd", PCD_BASEDIV16);
'
' Flag Value:               Meaning:
' PCD_BASE             24-bit, 768 x 512 pixels
' PCD_BASEDIV4         24-bit, 192 x 128 pixels
' PCD_BASEDIV16        24-bit. 384 x 256 pixels
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPCD(const char *filename, int flags FI_DEFAULT(PCD_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPCD Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPCD@8" (ByVal FileName As String, Optional ByVal flags As Long = PCD_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given PCX file into a FreeImage bitmap. If the file is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadPCX returns NULL.
'
' FIBITMAP *dib = FreeImage_LoadPCX("test.pcx");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPCX(const char *filename, int flags FI_DEFAULT(PCX_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPCX Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPCX@8" (ByVal FileName As String, Optional ByVal flags As Long = PCX_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given PNG (Portable Network Graphic) file into a FreeImage bitmap. If the file is loaded
' successfully, memory for it is allocated and a bitmap pointer is returned. If the file couldn't be loaded,
' FreeImage_LoadPNG returns NULL.
'
' FIBITMAP *dib = FreeImage_LoadPNG("test.png");
'
' Flag Value:               Meaning:
' PNG_IGNOREGAMMA      Avoid gamma correction
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPNG(const char *filename, int flags FI_DEFAULT(PNG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPNG Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPNG@8" (ByVal FileName As String, Optional ByVal flags As Long = PNG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads one of the PNM supported bitmap files into a FreeImage bitmap. If the file is loaded successfully,
' memory for it is allocated and a bitmap pointer is returned. If the file couldn't be loaded,
' FreeImage_LoadPNM returns NULL.
'
' PNM is a descriptive name for a collection of ASCII based bitmap types: PBM, PGM and PPM. PBM (Portable
' Bitmap format) is a 1-bit, black and white bitmap type, PGM (Portable Greymap format) is an 8-bit
' greyscale bitmap type and PPM (Portable Pixelmap format) is a 24-bit full colour bitmap.  FreeImage
' automatically detects if you are dealing with a PBM, PGM or PPM bitmap and will load it in the default bit
' depth of that particular image.
'
' FIBITMAP *dib = FreeImage_LoadPNM("test.ppm");
' FIBITMAP *dib = FreeImage_LoadPNM("test.pgm");
' FIBITMAP *dib = FreeImage_LoadPNM("test.pbm");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPNM(const char *filename, int flags FI_DEFAULT(PNM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPNM Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPNM@8" (ByVal FileName As String, Optional ByVal flags As Long = PNM_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads an Adobe Photoshop file (deprecated)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPSD(const char *filename, int flags FI_DEFAULT(PSD_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPSD Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPSD@8" (ByVal FileName As String, Optional ByVal flags As Long = PSD_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Sun rasterfile into a FreeImage bitmap. If the file is loaded successfully, memory for it
' is allocated and a bitmap pointer is returned. If the file couldn 't be loaded, FreeImage_LoadRAS returns
' NULL.
'
' FIBITMAP *dib = FreeImage_LoadRAS("test.ras");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadRAS(const char *filename, int flags FI_DEFAULT(RAS_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadRAS Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadRAS@8" (ByVal FileName As String, Optional ByVal flags As Long = RAS_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given TARGA file into a FreeImage bitmap. If the file is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the file couldn 't be loaded, FreeImage_LoadTARGA returns NULL.
'
' Flags defaults to TARGA_DEFAULT. The flags parameter can be passed the following values to change the
' behaviour of the load plugin:
'
' FIBITMAP *dib = FreeImage_LoadTARGA("test.tga");
' FIBITMAP *dib = FreeImage_LoadTARGA("test.tga", TARGA_LOAD_RGB888);
'
' Flag Value:               Meaning:
' TARGA_LOAD_RGB888    If the TARGA file is 16 or 32-bit, it is automatically converted to 24-bit.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadTARGA(const char *filename, int flags FI_DEFAULT(TARGA_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadTARGA Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadTARGA@8" (ByVal FileName As String, Optional ByVal flags As Long = TARGA_LOAD_RGB888) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given TIFF file into a FreeImage bitmap. If the file is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadTIFF returns NULL.
'
' FIBITMAP *dib = FreeImage_LoadTIFF("test.tiff");
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadTIFF(const char *filename, int flags FI_DEFAULT(TIFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadTIFF Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadTIFF@8" (ByVal FileName As String, Optional ByVal flags As Long = TIFF_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given WBMP or WAP file into a FreeImage bitmap. If the file is loaded successfully, memory for it
' is allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadWBMP returns
' NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadWBMP(const char *filename, int flags FI_DEFAULT(WBMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadWBMP Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadWBMP@8" (ByVal FileName As String, Optional ByVal flags As Long = WBMP_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given BMP file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadBMPFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadBMPFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(BMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadBMPFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadBMPFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = BMP_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads a Dr.Halo file from a handle (deprecated)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadCUTFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(CUT_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadCUTFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadCUTFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = CUT_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given ICO file into a FreeImage bitmap using the given FreeImageIO struct and fi_handle. If the
' file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadICOFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadICOFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(ICO_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadICOFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadICOFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = ICO_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given TIFF file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadTIFFFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadIFFFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(IFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadIFFFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadTIFFFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = TIFF_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given JPEG file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadJPEGFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadJPEGFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(JPEG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadJPEGFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadJPEGFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = JPEG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given KOALA file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadKOALAFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadKOALAFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(KOALA_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadKOALAFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadKOALAFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = KOALA_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Deluxe Paint LBM file into a FreeImage bitmap using the specified FreeImageIO struct and
' fi_handle. If the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned.
' If the file couldn't be loaded, FreeImage_LoadLBMFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadLBMFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(LBM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadLBMFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadLBMFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = LBM_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the Multi Network Graphics (MNG) or JPEG Network Graphics (JNG) file into a FreeImage bitmap using
' the specified FreeImageIO struct and fi_handle. If the file is loaded successfully, memory for it is
' allocated and a bitmap pointer is returned. If the file couldn't be loaded, FreeImage_LoadMNGFromHandle
' returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadMNGFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(MNG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadMNGFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadMNGFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = MNG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Kodak PhotoCD file into a FreeImage bitmap using the specified FreeImageIO struct and
' fi_handle. If the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned.
' If the file couldn't be loaded, FreeImage_LoadPCDFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPCDFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PCD_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPCDFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPCDFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PCD_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given PCX file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadPCXFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPCXFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PCX_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPCXFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPCXFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PCX_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given PNG file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadPNGFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPNGFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PNG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPNGFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPNGFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PNG_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given PBM, PGM or PPM file into a FreeImage bitmap using the specified FreeImageIO struct and
' fi_handle. If the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned.
' If the file couldn't be loaded, FreeImage_LoadPNMFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPNMFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PNM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPNMFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPNMFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PNM_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads an Adobe Photoshop Document from a handle (deprecated)
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadPSDFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PSD_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadPSDFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadPSDFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PSD_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given Sun Rasterfile into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle.
' If the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadRASFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadRASFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(RAS_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadRASFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadRASFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = RAS_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given TARGA file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadTARGAFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadTARGAFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(TARGA_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadTARGAFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadTARGAFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = TARGA_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given TIFF file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadTIFFFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadTIFFFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(TIFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadTIFFFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadTIFFFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = TIFF_DEFAULT) As Long

'-------------------------------------------------------------------------------------------------------------
' Loads the given WBMP file into a FreeImage bitmap using the specified FreeImageIO struct and fi_handle. If
' the file is loaded successfully, memory for it is allocated and a bitmap pointer is returned. If the file
' couldn't be loaded, FreeImage_LoadWBMPFromHandle returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_LoadWBMPFromHandle(FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(WBMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_LoadWBMPFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_LoadWBMPFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = WBMP_DEFAULT) As Long

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Bitmap save routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a Windows Bitmap file. The BMP file is always saved in the Windows format.
' No compression is used.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
' if (dib != NULL) {
'   FreeImage_SaveBMP("saved.bmp");
'   FreeImage_Free(dib);
' }
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveBMP(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(BMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveBMP Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveBMP@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = BMP_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a JPEG file.
' NOTE: Only 24-bit bitmaps can be saved as JPEG. Bitmaps in OTHER bit depths will have to be converted to 24-bit.
' NOTE: If you are saving a JPEG image, you can pass the quality of the JPEG image (1 to 100) as the "flags" parameter
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
' if (dib != NULL) {
'   FreeImage_SaveJPEG("saved.jpg");
'   FreeImage_Free(dib);
' }
'
' Flag Value:               Meaning:
' JPEG_FAST
' JPEG_ACCURATE
' JPEG_QUALITYSUPERB
' JPEG_QUALITYGOOD
' JPEG_QUALITYNORMAL
' JPEG_QUALITYAVERAGE
' JPEG_QUALITYBAD
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveJPEG(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(JPEG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveJPEG Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveJPEG@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = JPEG_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a PNG file.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
' if (dib != NULL) {
'   FreeImage_SavePNG("saved.png");
'   FreeImage_Free(dib);
' }
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SavePNG(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT( PNG_DEFAULT ));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SavePNG Lib "FREEIMAGE.DLL" Alias "_FreeImage_SavePNG@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = PNG_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a PNM file. PNM is a descriptive name for a collection of ASCII based bitmap
' types: PBM, PGM and PPM. If the bitmap has a bit depth of 1, the file is saved as a PBM file. If the
' bitmap has a bit depth of 8, the file is saved as a PGM file. If the bitmap has a bit depth of 24, the
' file is saved as a PPM file. Other bit depths are not supported.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
'
' if (dib != NULL) {
'   switch(FreeImage_GetBPP(dib)) {
'     Case 1:
'       FreeImage_SavePNM("saved.pbm");
'       break;
'     Case 8:
'       FreeImage_SavePNM("saved.pgm");
'       break;
'     Case 24:
'       FreeImage_SavePNM("saved.ppm");
'       break;
'   }
'   FreeImage_Free(dib);
' }
'
' Flag Value:               Meaning:
' PNM_SAVE_RAW         If set the writer saves in RAW format (i.e. P4, P5 or P6)
' PNM_SAVE_ASCII       If set the writer saves in ASCII format (i.e. P1, P2 or P3)
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SavePNM(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(PNM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SavePNM Lib "FREEIMAGE.DLL" Alias "_FreeImage_SavePNN@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = PNM_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a TIFF file.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
' if (dib != NULL) {
'   FreeImage_SaveTIFF("saved.tiff");
'   FreeImage_Free(dib);
' }
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveTIFF(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(TIFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveTIFF Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveTIFF@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = TIFF_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a WBMP file.
'
' FIBITMAP *dib = FreeImage_LoadBMP("test.bmp");
' if (dib != NULL) {
'   FreeImage_SaveWBMP("saved.wbmp");
'   FreeImage_Free(dib);
' }
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveWBMP(FIBITMAP *dib, const char *filename, int flags FI_DEFAULT(WBMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveWBMP Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveWBMP@12" (ByVal DIB As Long, ByVal FileName As String, Optional ByVal flags As Long = WBMP_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a Windows Bitmap file. The BMP file is always saved in the Windows format.
' No compression is used.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveBMPToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(BMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveBMPToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveBMPToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = BMP_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a JPEG file. Only 24-bit bitmaps can be saved as JPEG. Bitmaps in bit depths will
' have to be converted.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveJPEGToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(JPEG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveJPEGToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveJPEGToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = JPEG_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a PNG file.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SavePNGToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PNG_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SavePNGToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SavePNGToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PNG_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a PNM file. PNM is a descriptive name for a collection of ASCII based bitmap
' types: PBM, PGM and PPM. If the bitmap has a bitdepth of 1, the file is saved as a PBM file. If the bitmap
' has a bitdepth of 8, the file is saved as a PGM file. If the bitmap has a bitdepth of 24, the file is saved
' as a PPM file. Other bitdepths are not supported.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SavePNMToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(PNM_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SavePNMToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SavePNMToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = PNM_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a TIFF file.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveTIFFToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(TIFF_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveTIFFToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveTIFFToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = TIFF_DEFAULT) As BOOL

'-------------------------------------------------------------------------------------------------------------
' Saves the FreeImage DIB to a WBMP file.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_SaveWBMPToHandle(FIBITMAP *dib, FreeImageIO *io, fi_handle handle, int flags FI_DEFAULT(WBMP_DEFAULT));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_SaveWBMPToHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_SaveWBMPToHandle@16" (ByVal DIB As Long, ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, Optional ByVal flags As Long = WBMP_DEFAULT) As BOOL

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Filetype request routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Investigates the bitmap data from the given bitmap and returns the FreeImage Format ID for it.
' FreeImage_GetFileType will read the first size bytes of the file but not more than 16. FreeImage_GetFileType
' will return one of the following values:
'
' FIF_UNKNOWN = -1  = Unknown format
' FIF_BMP = 0       = Windows or OS/2 Bitmap File (*.BMP)
' FIF_ICO           = Windows Icon (*.ICO)
' FIF_JPEG          = Independent JPEG Group (*.JPG)
' FIF_JNG           = JPEG Network Graphics (*.JNG)
' FIF_KOALA         = Commodore 64 Koala format (*.KOA)
' FIF_LBM           = Amiga IFF (*.IFF, *.LBM)
' FIF_MNG           = Multiple Network Graphics (*.MNG)
' FIF_PBM           = Portable Bitmap (ASCII) (*.PBM)
' FIF_PBMRAW        = Portable Bitmap (BINARY) (*.PBM / *.PBMRAW)
' FIF_PCD           = Kodak PhotoCD (*.PCD)
' FIF_PCX           = PCX bitmap format (*.PCX)
' FIF_PGM           = Portable Graymap (ASCII) (*.PGM)
' FIF_PGMRAW        = Portable Graymap (BINARY) (*.PGM / *.PGMRAW)
' FIF_PNG           = Portable Network Graphics (*.PNG)
' FIF_PPM           = Portable Pixelmap (ASCII) (*.PPM)
' FIF_PPMRAW        = Portable Pixelmap (BINARY) (*.PPM / *.PPMRAW)
' FIF_RAS           = Sun Rasterfile (*.RAS)
' FIF_TARGA         = Targa files (*.TGA)
' FIF_TIFF          = Tagged Image File Format (*.TIFF)
' FIF_WBMP          = Wireless Bitmap (*.WBMP)
' FIF_PSD           = Adobe Photoshop (*.PSD)
' FIF_CUT           = Dr. Halo (*.CUT)
' FIF_IFF = FIF_LBM = Amiga IFF (*.IFF, *.LBM)
' FIF_XBM           = X11 Bitmap Format (*.XBM)
'
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFileType(const char *filename, int size);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFileType Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFileType@8" (ByVal FileName As String, Optional ByVal Size As Long = 16) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Investigates the bitmap data from the given bitmap and returns the FreeImage Format ID for it.
'
' Code example:
'
' unsigned
' _ReadProc(FIBITMAP *buffer, unsigned s, unsigned c, fi_handle handle) {
'   return fread(buffer, s, c, (FILE *)handle);
' }
' unsigned
' _WriteProc(FIBITMAP *buffer, unsigned s, unsigned c, fi_handle handle){
'   return fwrite(buffer, s, c, (FILE *)handle);
' }
' int
' _SeekProc(fi_handle handle, long offset, int origin) {
'   return fseek((FILE *)handle, offset, origin);
' }
' long
' _TellProc(fi_handle handle) {
'   return ftell((FILE *)handle);
' }
' FreeImageIO io;
' io.read_proc  = _ReadProc;
' io.write_proc = _WriteProc;
' io.seek_proc  = _SeekProc;
' io.tell_proc  = _TellProc;
' FILE *file = fopen("test.bmp", "rb");
' FREE_IMAGE_TYPE type;
' type = FreeImage_GetFileTypeFromHandle(&io, (fi_handle)file, 16);
'
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFileTypeFromHandle(FreeImageIO *io, fi_handle handle, int size);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFileTypeFromHandle Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFileTypeFromHandle@12" (ByRef IO_Callback_Functions As FreeImageIO, ByVal Handle As Long, ByVal Size As Long) As FREE_IMAGE_FORMAT

'-------------------------------------------------------------------------------------------------------------
' Takes the specified FreeImage Format ID and returns back the string representation of it.
'-------------------------------------------------------------------------------------------------------------
'DLL_API const char * DLL_CALLCONV FreeImage_GetFileTypeFromFormat(FREE_IMAGE_FORMAT fif);  // this function is deprecated
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFileTypeFromFormat Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFileTypeFromFormat@4" (ByVal ImageFormat As FREE_IMAGE_FORMAT) As String

'-------------------------------------------------------------------------------------------------------------
' This function is deprecated. FreeImage_GetFIFFromFilename has replaced its functionality.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_FORMAT DLL_CALLCONV FreeImage_GetFileTypeFromExt(const char *filename); // this function is deprecated
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetFileTypeFromExt Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetFileTypeFromExt@4" (ByVal FileName As String) As FREE_IMAGE_FORMAT

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' FreeImage info routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Returns the red bitmask for the bitmap. If the bitmap is palletised 0 is returned.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetRedMask(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetRedMask Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetRedMask@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the green bitmask for the bitmap. If the bitmap is palletised 0 is returned.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetGreenMask(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetGreenMask Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetGreenMask@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the blue bitmask for the bitmap. If the bitmap is palletised 0 is returned.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetBlueMask(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetBlueMask Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetBlueMask@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the number of transparent colours stored into the transparency colour table of the given bitmap.
' Every palletised bitmap includes a transparency table containing up to 256 alpha values.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetTransparencyCount(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetTransparencyCount Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetTransparencyCount@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the transparency table assigned to this bitmap. If the bitmap doesn't contain a transparency table
' NULL is returned.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BYTE * DLL_CALLCONV FreeImage_GetTransparencyTable(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetTransparencyTable Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetTransparencyTable@4" (ByVal DIB As Long) As Long 'RETURNS BYTE ARRAY... THIS DECLARATION WILL RETURN A POINTER TO THAT ARRAY

'-------------------------------------------------------------------------------------------------------------
' Flags a bitmap as having transparent areas.  Only used in palettized bitmaps.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_SetTransparent(FIBITMAP *dib, BOOL enabled);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_SetTransparent Lib "FREEIMAGE.DLL" Alias "_FreeImage_SetTransparent@8" (ByVal DIB As Long, ByVal Enabled As BOOL)

'-------------------------------------------------------------------------------------------------------------
' Assigns a new transparency table to the bitmap. A transparency table consists of up to 256 bytes in an
' array, where a value of 0xFF in that table stands for completely opaque and 0x00 stands for completely
' invisible. The transparency table is used in TrollTech Qt to draw transparent bitmaps on a widget.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_SetTransparencyTable(FIBITMAP *dib, BYTE *table, int count);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_SetTransparencyTable Lib "FREEIMAGE.DLL" Alias "_FreeImage_SetTransparencyTable@12" (ByVal DIB As Long, ByRef Table() As Byte, ByVal Count As Long)

'-------------------------------------------------------------------------------------------------------------
' Returns the transparency status flag.  Only used in palettized bitmaps.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BOOL DLL_CALLCONV FreeImage_IsTransparent(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_IsTransparent Lib "FREEIMAGE.DLL" Alias "_FreeImage_IsTransparent@4" (ByVal DIB As Long) As BOOL

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' DIB info routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Retrieves the number of colours used in the bitmap. If the bitmap is non-palletised, 0 is returned.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetColorsUsed(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetColorsUsed Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetColorsUsed@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the bitmap bits.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BYTE *DLL_CALLCONV FreeImage_GetBits(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetBits Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetBits@4" (ByVal DIB As Long) As Long 'RETURNS BYTE ARRAY... THIS DECLARATION WILL RETURN A POINTER TO THAT ARRAY

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the bitmap bits on the given column and row.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BYTE *DLL_CALLCONV FreeImage_GetBitsRowCol(FIBITMAP *dib, int col, int row);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetBitsRowCol Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetBitsRowCol@12" (ByVal DIB As Long, ByVal Col As Long, ByVal Row As Long) As Long  'RETURNS BYTE ARRAY... THIS DECLARATION WILL RETURN A POINTER TO THAT ARRAY

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the beginning of the bits of the given scanline.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BYTE *DLL_CALLCONV FreeImage_GetScanLine(FIBITMAP *dib, int scanline);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetScanLine Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetScanLine@8" (ByVal DIB As Long, ByVal ScanLine As Long) As Long  'RETURNS BYTE ARRAY... THIS DECLARATION WILL RETURN A POINTER TO THAT ARRAY

'-------------------------------------------------------------------------------------------------------------
' Returns the bitdepth of the bitmap.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetBPP(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetBPP Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetBPP@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the width of the bitmap in pixels.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetWidth(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetWidth Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetWidth@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the height of the bitmap in pixels.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetHeight(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetHeight Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetHeight@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the width of a bitmap in BYTES (not pixels)
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetLine(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetLine Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetLine@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the width of the bitmap in BYTES (instead of pixels) rounded to the nearest DWORD.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetPitch(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetPitch Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetPitch@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the size of the bitmap in bytes. The size of the bitmap is the BITMAPINFOHEADER + the size of the
' palette + the size of the bitmap data.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetDIBSize(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetDIBSize Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetDIBSize@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the bitmap's palette.  If the bitmap doesn't have a palette, FreeImage_GetPalette
' returns NULL.
'-------------------------------------------------------------------------------------------------------------
'DLL_API RGBQUAD *DLL_CALLCONV FreeImage_GetPalette(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetPalette Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetPalette@4" (ByVal DIB As Long) As Long 'RETURNS RGBQUAD ARRAY... THIS DECLARATION WILL RETURN A POINTER TO THAT ARRAY

'-------------------------------------------------------------------------------------------------------------
' Returns the horizontal resolution, in pixels-per-meter, of the target device for the bitmap.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetDotsPerMeterX(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetDotsPerMeterX Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetDotsPerMeterX@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns the vertical resolution, in pixels-per-meter, of the target device for the bitmap.
'-------------------------------------------------------------------------------------------------------------
'DLL_API unsigned DLL_CALLCONV FreeImage_GetDotsPerMeterY(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetDotsPerMeterY Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetDotsPerMeterY@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the bitmap's BITMAPINFOHEADER.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BITMAPINFOHEADER *DLL_CALLCONV FreeImage_GetInfoHeader(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetInfoHeader Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetInfoHeader@4" (ByVal DIB As Long) As Long 'BITMAPINFOHEADER

'-------------------------------------------------------------------------------------------------------------
' Returns a pointer to the bitmap's BITMAPINFO header.
'-------------------------------------------------------------------------------------------------------------
'DLL_API BITMAPINFO *DLL_CALLCONV FreeImage_GetInfo(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetInfo Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetInfo@4" (ByVal DIB As Long) As Long 'BITMAPINFO

'-------------------------------------------------------------------------------------------------------------
' Investigates the colour type of the bitmap and returns one of the following values:
'
' Value            Description
' FIC_MINISWHITE   1-bit bitmap. The min value is white
' FIC_MINISBLACK   1-bit bitmap. The min value is black 8-bit grayscale. The min value is black
' FIC_RGB          24/32-bit RGB
' FIC_PALETTE      1/4/8-bit palletised
'-------------------------------------------------------------------------------------------------------------
'DLL_API FREE_IMAGE_COLOR_TYPE DLL_CALLCONV FreeImage_GetColorType(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetColorType Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetColorType@4" (ByVal DIB As Long) As FREE_IMAGE_COLOR_TYPE

'_____________________________________________________________________________________________________________
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Conversion routines
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine1To8(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine1To8 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine1To8@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine4To8(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine4To8 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine4To8@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To8_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To8_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To8_555@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To8_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To8_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To8_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine24To8(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine24To8 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine24To8@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine32To8(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine32To8 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine32To8@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine1To16_555(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine1To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine1To16_555@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine4To16_555(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine4To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine4To16_555@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine8To16_555(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine8To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine8To16_555@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16_565_To16_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16_565_To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16_565_To16_555@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine24To16_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine24To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine24To16_555@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine32To16_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine32To16_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine32To16_555@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine1To16_565(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine1To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine1To16_565@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine4To16_565(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine4To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine4To16_565@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine8To16_565(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine8To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine8To16_565@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16_555_To16_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16_555_To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16_555_To16_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine24To16_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine24To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine24To16_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine32To16_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine32To16_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine32To16_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine1To24(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine1To24 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine1To24@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine4To24(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine4To24 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine4To24@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine8To24(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine8To24 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine8To24@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To24_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To24_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To24_555@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To24_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To24_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To24_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine32To24(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine32To24 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine32To24@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine1To32(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine1To32 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine1To32@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine4To32(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine4To32 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine4To32@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine8To32(BYTE *target, BYTE *source, int width_in_pixels, RGBQUAD *palette);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine8To32 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine8To32@16" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long, ByRef Palette() As RGBQUAD)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To32_555(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To32_555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To32_555@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine16To32_565(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine16To32_565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine16To32_565@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap scanline from the source color depth (bits per pixel) to the target color depth (BPP)
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertLine24To32(BYTE *target, BYTE *source, int width_in_pixels);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertLine24To32 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertLine24To32@12" (ByRef Target() As Byte, ByRef Source() As Byte, ByVal width_in_pixels As Long)

'-------------------------------------------------------------------------------------------------------------
' Converts the given bitmap to 8 bits. If the bitmap is 24 or 32-bit RGB, the colour values are converted to
' greyscale.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertTo8Bits(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertTo8Bits Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertTo8Bits@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Converts the given bitmap to 16 bits. The resulting bitmap has a layout of 5 bits red, 5 bits green,
' 5 bits red and 1 unused bit.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertTo16Bits555(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertTo16Bits555 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertTo16Bits555@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Converts the given bitmap to 16 bits. The resulting bitmap has a layout of 5 bits red, 6 bits green and 5
' bits red.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertTo16Bits565(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertTo16Bits565 Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertTo16Bits565@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Converts the given bitmap to 24 bits.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertTo24Bits(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertTo24Bits Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertTo24Bits@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Converts the given bitmap to 32 bits.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertTo32Bits(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertTo32Bits Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertTo32Bits@4" (ByVal DIB As Long) As Long

'-------------------------------------------------------------------------------------------------------------
' Quantizes a full colour 24-bit bitmap to a palletised 8-bit bitmap. The quantize parameter specifies
' which colour reduction algorithm should be used.
'
' Parameter     Description
' FIQ_WUQUANT   Wu's color quantization algorithm
' FIQ_NNQUANT   NeuQuant quantization algorithm
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ColorQuantize(FIBITMAP *dib, FREE_IMAGE_QUANTIZE quantize);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ColorQuantize Lib "FREEIMAGE.DLL" Alias "_FreeImage_ColorQuantize@8" (ByVal DIB As Long, ByVal Quantize As FREE_IMAGE_QUANTIZE) As Long

'-------------------------------------------------------------------------------------------------------------
' Grabs a raw piece of memory and converts it into a FreeImage bitmap
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_ConvertFromRawBits(BYTE *bits, int width, int height, int pitch, unsigned bpp, unsigned red_mask, unsigned green_mask, unsigned blue_mask, BOOL topdown FI_DEFAULT(FALSE));
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_ConvertFromRawBits Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertFromRawBits@36" (ByRef Bits() As Byte, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal Red_Mask As Long = 0, Optional ByVal Green_Mask As Long = 0, Optional ByVal Blue_Mask As Long = 0, Optional ByVal TopDown As BOOL = FALSE_) As Long

'-------------------------------------------------------------------------------------------------------------
' Writes a freeimage bitmap to a raw piece of memory in a certain format.  Usefull for example to load a
' bitmap into a DirectX surface.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_ConvertToRawBits(BYTE *bits, FIBITMAP *dib, int pitch, unsigned bpp, unsigned red_mask, unsigned green_mask, unsigned blue_mask, BOOL topdown FI_DEFAULT(FALSE));
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_ConvertToRawBits Lib "FREEIMAGE.DLL" Alias "_FreeImage_ConvertToRawBits@32" (ByRef Bits() As Byte, ByVal DIB As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal Red_Mask As Long, ByVal Green_Mask As Long, ByVal Blue_Mask As Long, Optional ByVal TopDown As BOOL = FALSE_)

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap to 1-bit monochrome bitmap using a threshold T between [0..255].  The function first
' converts the bitmap to a 8-bit greyscale bitmap. Then, any brightness level that is less than T is set to
' zero, otherwise to 1.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_Threshold(FIBITMAP *dib, BYTE T);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Threshold Lib "FREEIMAGE.DLL" Alias "_FreeImage_Threshold@8" (ByVal DIB As Long, Optional ByVal bytThreshold As Byte = 128) As Long
Public Declare Function FreeImage_Monochrome Lib "FREEIMAGE.DLL" Alias "_FreeImage_Threshold@8" (ByVal DIB As Long, Optional ByVal bytThreshold As Byte = 128) As Long

'-------------------------------------------------------------------------------------------------------------
' Converts a bitmap to 1-bit monochrome bitmap using a dithering algorithm. The algorithm parameter specifies
' the dithering algorithm to be used.
'
' The function first converts the bitmap to a 8-bit greyscale bitmap. Then, the bitmap is dithered using one
' of the following algorithm:
'
' FID_FS            Floyd & Steinberg error diffusion algorithm
' FID_BAYER4x4      Bayer ordered dispersed dot dithering (order 2 – 4x4 - dithering matrix)
' FID_BAYER8x8      Bayer ordered dispersed dot dithering (order 3 – 8x8 - dithering matrix)
' FID_CLUSTER6x6    Ordered clustered dot dithering (order 3 - 6x6 matrix)
' FID_CLUSTER8x8    Ordered clustered dot dithering (order 4 - 8x8 matrix)
' FID_CLUSTER16x16  Ordered clustered dot dithering (order 8 - 16x16 matrix)
'
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIBITMAP *DLL_CALLCONV FreeImage_Dither(FIBITMAP *dib, FREE_IMAGE_DITHER algorithm);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_Dither Lib "FREEIMAGE.DLL" Alias "_FreeImage_Dither@8" (ByVal DIB As Long, ByVal Algorithm As FREE_IMAGE_DITHER) As Long

'-------------------------------------------------------------------------------------------------------------
' Retrieve the a pointer to the FIICCPROFILE data of the bitmap. This function can also be called safely,
' when the original format does not support profiles.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIICCPROFILE *DLL_CALLCONV FreeImage_GetICCProfile(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_GetICCProfile Lib "FREEIMAGE.DLL" Alias "_FreeImage_GetICCProfile@4" (ByVal DIB As Long) As FIICCPROFILE

'-------------------------------------------------------------------------------------------------------------
' This functions destroys an FIICCPROFILE previously created by FreeImage_CreateICCProfile. After this call
' the bitmap will contain no profile information. This function should be called to ensure that a stored
' bitmap will not contain any profile information.
'-------------------------------------------------------------------------------------------------------------
'DLL_API void DLL_CALLCONV FreeImage_DestroyICCProfile(FIBITMAP *dib);
'-------------------------------------------------------------------------------------------------------------
Public Declare Sub FreeImage_DestroyICCProfile Lib "FREEIMAGE.DLL" Alias "_FreeImage_DestroyICCProfile@4" (ByVal DIB As Long)

'-------------------------------------------------------------------------------------------------------------
' Create a new FIICCPROFILE block from ICC profile data previously read from a file or build by a colour
' management system. The profile data are attached to the bitmap. The function returns a pointer to the
' FIICCPROFILE structure created.
'-------------------------------------------------------------------------------------------------------------
'DLL_API FIICCPROFILE *DLL_CALLCONV FreeImage_CreateICCProfile(FIBITMAP *dib, void *data, long size);
'-------------------------------------------------------------------------------------------------------------
Public Declare Function FreeImage_CreateICCProfile Lib "FREEIMAGE.DLL" Alias "_FreeImage_CreateICCProfile@12" (ByVal DIB, ByRef Data() As Byte, ByVal Size As Long) As FIICCPROFILE



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                          WIN32  API  DECLARATIONS                       XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' CreateDIBitmap.fdwInit
Public Const CBM_INIT As Long = &H4 ' If this flag is set, the operating system uses the data pointed to by the lpbInit and lpbmi parameters to initialize the bitmap’s bits.If this flag is clear, the data pointed to by those parameters is not used.

' SetDIBitsToDevice.ColorUse
Public Const DIB_RGB_COLORS = 0 ' Color table in RGBs
Public Const DIB_PAL_COLORS = 1 ' Color table in palette indices

' SetStretchBltMode.StretchMode
Public Const BLACKONWHITE        As Long = 1 ' Performs a Boolean AND operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves black pixels at the expense of white pixels.
Public Const WHITEONBLACK        As Long = 2 ' Performs a Boolean OR operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves white pixels at the expense of black pixels.
Public Const COLORONCOLOR        As Long = 3 ' Deletes the pixels. This mode deletes all eliminated lines of pixels without trying to preserve their information.
Public Const HALFTONE            As Long = 4 ' Maps pixels from the source rectangle into blocks of pixels in the destination rectangle. The average color over the destination block of pixels approximates the color of the source pixels. After setting the HALFTONE stretching mode, an application must call the SetBrushOrgEx function to set the brush origin. If it fails to do so, brush misalignment occurs.
Public Const STRETCH_ANDSCANS    As Long = BLACKONWHITE ' Same as BLACKONWHITE.
Public Const STRETCH_DELETESCANS As Long = COLORONCOLOR ' Same as COLORONCOLOR.
Public Const STRETCH_HALFTONE    As Long = HALFTONE     ' Same as HALFTONE.
Public Const STRETCH_ORSCANS     As Long = WHITEONBLACK ' Same as WHITEONBLACK.

' GetObjectType (Return)
Public Const OBJ_PEN         As Long = 1  ' Pen
Public Const OBJ_BRUSH       As Long = 2  ' Brush
Public Const OBJ_DC          As Long = 3  ' Device context
Public Const OBJ_METADC      As Long = 4  ' Metafile device context
Public Const OBJ_PAL         As Long = 5  ' Palette
Public Const OBJ_FONT        As Long = 6  ' Font
Public Const OBJ_BITMAP      As Long = 7  ' BITMAP
Public Const OBJ_REGION      As Long = 8  ' Region
Public Const OBJ_METAFILE    As Long = 9  ' metafile
Public Const OBJ_MEMDC       As Long = 10 ' Memory device context
Public Const OBJ_EXTPEN      As Long = 11 ' Extended pen
Public Const OBJ_ENHMETADC   As Long = 12 ' Enhanced metafile device context
Public Const OBJ_ENHMETAFILE As Long = 13 ' Enhanced metafile
Public Const OBJ_COLORSPACE  As Long = 14 ' Color Space

' Constants - CopyImage.fuFlags
Public Const LR_COPYDELETEORG = &H8       ' Deletes the original image after creating the copy.
Public Const LR_COPYFROMRESOURCE = &H4000 ' Tries to reload an icon or cursor resource from the original resource file rather than simply copying the current image. This is useful for creating a different-sized copy when the resource file contains multiple sizes of the resource. Without this flag, CopyImage stretches the original image to the new size. If this flag is set, CopyImage uses the size in the resource file closest to the desired size.  This will succeed only if hImage was loaded by LoadIcon or LoadCursor, or by LoadImage with the LR_SHARED flag.
Public Const LR_COPYRETURNORG = &H4       ' Returns the original hImage if it satisfies the criteria for the copy—that is, correct dimensions and color depth—in which case the LR_COPYDELETEORG flag is ignored. If this flag is not specified, a new object is always created.
Public Const LR_CREATEDIBSECTION = &H2000 ' If this is set and a new bitmap is created, the bitmap is created as a DIB section. Otherwise, the bitmap image is created as a device-dependent bitmap. This flag is only valid if uType is IMAGE_BITMAP.
Public Const LR_MONOCHROME = &H1          ' Creates a new monochrome image.

' Constants - BITMAP.bmType & CopyImage.uType
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_CURSOR = 1
Public Const IMAGE_ICON = 2
Public Const IMAGE_ENHMETAFILE = 3

' Win32 API Sub/Function Declarations
Public Declare Sub CopyMemory Lib "KERNEL32.DLL" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function BitBlt Lib "GDI32.DLL" (ByVal Dest_hDC As Long, ByVal Dest_X As Long, ByVal Dest_Y As Long, ByVal Dest_Width As Long, ByVal Dest_Height As Long, ByVal Src_hDC As Long, ByVal Src_X As Long, ByVal Src_Y As Long, ByVal RasterOperation As RasterOpConstants) As BOOL
Public Declare Function CopyImage Lib "USER32.DLL" (ByVal hImage As Long, ByVal uType As Long, ByVal OutputWidth As Long, ByVal OutputHeight As Long, ByVal fuFlags As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBitmap Lib "GDI32.DLL" (ByVal hdc As Long, ByRef lpBITMAPINFOHEADER As Any, ByVal BitFlags As Long, ByRef lpBits As Any, ByRef lpBitmapInfo As Any, ByVal fuUsage As Long) As Long
Public Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hdc As Long) As BOOL
'Public Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hObject As Long) As BOOL
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Public Declare Function GetDIBits Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hBITMAP As Long, ByVal StartingScanLine As Long, ByVal ScanLineCount As Long, ByRef lpBits As Any, ByRef lpBitmapInfo As Any, ByVal ColorUsage As Long) As Long
Public Declare Function GetObjectAPI Lib "GDI32.DLL" Alias "GetObjectA" (ByVal hObject As Long, ByVal BufferLength As Long, ByRef ObjectInfo As Any) As Long
Public Declare Function GetObjectType Lib "GDI32.DLL" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "GDI32.DLL" (ByVal hdc As Long, ByVal Xpos As Long, ByVal nYPos As Long) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "GDI32.DLL" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "GDI32.DLL" (ByVal Dest_hDC As Long, ByVal Dest_X As Long, ByVal Dest_Y As Long, ByVal Src_Width As Long, ByVal Src_Height As Long, ByVal Src_X As Long, ByVal Src_Y As Long, ByVal StartingScanLine As Long, ByVal ScanLineCount As Long, ByRef lpBits As Any, ByRef lpBMI As Any, ByVal ColorUse As Long) As Long
Public Declare Function SetPixel Lib "GDI32.DLL" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetStretchBltMode Lib "GDI32.DLL" (ByVal hdc As Long, ByVal StretchMode As Long) As Long
Public Declare Function StretchBlt Lib "GDI32.DLL" (ByVal Dest_hDC As Long, ByVal Dest_X As Long, ByVal Dest_Y As Long, ByVal Dest_Width As Long, ByVal Dest_Height As Long, ByVal Src_hDC As Long, ByVal Src_X As Long, ByVal Src_Y As Long, ByVal Src_Width As Long, ByVal Src_Height As Long, ByVal RasterOperation As RasterOpConstants) As BOOL
Public Declare Function lstrcpy Lib "KERNEL32.DLL" (ByVal lpString1 As String, ByRef lpString2 As Any) As Long
Public Declare Function lstrlen Lib "KERNEL32.DLL" (ByRef lpString As Any) As Long


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                    SUPPORTING FUNCTION  DECLARATIONS                    XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'-------------------------------------------------------------------------------------------------------------
' Callback function that receives status/error messages generated inside the DLL. If a bitmap load action
' fails for example, this function receives why it failed.
'-------------------------------------------------------------------------------------------------------------
'typedef void (*FreeImage_OutputMessageFunction)(FREE_IMAGE_FORMAT fif, const char *msg);
'-------------------------------------------------------------------------------------------------------------
Public Sub FI_OutputMessageProc(ByVal ImageFormat As FREE_IMAGE_FORMAT, ByRef strMessage As Byte)
   
   Dim lngLen As Long
   
   ' Get the length of the message string
   lngLen = lstrlen(strMessage)
   If lngLen > 0 Then
      
      ' Allocate space enough for the string
      LastError_ErrDesc = String(lngLen, Chr(0))
      
      ' Copy the string into the buffer we just created
      If lstrcpy(LastError_ErrDesc, strMessage) = 0 Then LastError_ErrDesc = ""
      
   End If
   
End Sub

'typedef unsigned (DLL_CALLCONV *FI_ReadProc) (void *buffer, unsigned size, unsigned count, fi_handle handle);
'Public Function FI_ReadProc(ByVal Buffer As Long, ByVal Size As Long, ByVal Count As Long, ByVal Handle As Long) As Long
  
  'return fread(buffer, s, c, (FILE *)handle);
  
'End Function

'typedef unsigned (DLL_CALLCONV *FI_WriteProc) (void *buffer, unsigned size, unsigned count, fi_handle handle);
'Public Function FI_WriteProc(ByVal Buffer As Long, ByVal Size As Long, ByVal Count As Long, ByVal Handle As Long) As Long
  
  'return fwrite(buffer, s, c, (FILE *)handle);
  
'End Function

'typedef int (DLL_CALLCONV *FI_SeekProc) (fi_handle handle, long offset, int origin);
'Public Function FI_SeekProc(ByVal Handle As Long, ByVal Offset As Long, ByVal Origin As Long) As Long
  
  'return fseek((FILE *)handle, offset, origin);
  
'End Function

'typedef long (DLL_CALLCONV *FI_TellProc) (fi_handle handle);
'Public Function FI_TellProc(ByVal Handle As Long) As Long
  
  'return ftell((FILE *)handle);
  
'End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX                       CUSTOM  FUNCTION  DECLARATIONS                    XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'=============================================================================================================
' FreeImage_Initialise
'
' This is a replacement function that both initialises the FreeImage.dll library and sets up the CALLBACK for
' it so we can receive error messages sent to us by FreeImage.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' None
'=============================================================================================================
Public Sub FreeImage_Initialise(Optional ByVal blnLoadPlugins As Boolean = True)
   
   ' Initialise the FreeImage.dll library
   If blnLoadPlugins = True Then
      Call FreeImage_Init(FALSE_)
   Else
      Call FreeImage_Init(TRUE_)
   End If
   
   ' Set the CALLBACK for FreeImage so we can trap error messages being sent back to us
   Call FreeImage_SetOutputMessage(AddressOf FI_OutputMessageProc)
   
End Sub


'=============================================================================================================
' FreeImage_GetVersion
'
' This is a replacement function that makes the "FreeImage_GetVersion" function friendly to Visual Basic.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' A string containing the version of the FreeImage library.
'=============================================================================================================
Public Function FreeImage_GetVersion() As String
   
   Dim lngLen As Long
   Dim lngVer As Long
   
   ' Get a pointer to the version string
   lngVer = FreeImage_GetVer
   If lngVer <> 0 Then
      
      ' Get the length of the string
      lngLen = lstrlen(ByVal lngVer)
      If lngLen > 0 Then
         
         ' Allocate space enough for the string
         FreeImage_GetVersion = String(lngLen, Chr(0))
         
         ' Copy the string from a POINTER
         If lstrcpy(FreeImage_GetVersion, ByVal lngVer) = 0 Then FreeImage_GetVersion = ""
         
      End If
   End If
   
End Function

'=============================================================================================================
' FreeImage_GetCopyrightMessage
'
' This is a replacement function that makes the "FreeImage_GetCopyrightMessage" function friendly to
' Visual Basic.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' A string containing a short copyright message you can include in your program.
'=============================================================================================================
Public Function FreeImage_GetCopyrightMessage() As String
   
   Dim lngLen As Long
   Dim lngVer As Long
   
   ' Get a pointer to the copyright string
   lngVer = FreeImage_GetCopyrightMsg
   If lngVer <> 0 Then
      
      ' Get the length of the string
      lngLen = lstrlen(ByVal lngVer)
      If lngLen > 0 Then
         
         ' Allocate space enough for the string
         FreeImage_GetCopyrightMessage = String(lngLen, Chr(0))
         
         ' Copy the string from a POINTER
         If lstrcpy(FreeImage_GetCopyrightMessage, ByVal lngVer) = 0 Then FreeImage_GetCopyrightMessage = ""
         
      End If
   End If
   
End Function

'=============================================================================================================
' FreeImage_GetLastError
'
' This function returns the most recent message returned to the "FI_OutputMessageProc" CALLBACK function.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' If an error occured which was returned via the CALLBACK funciton, this returns the error description
' If no error has occured, this returns a blank string
'=============================================================================================================
Public Function FreeImage_GetLastError() As String
   
   FreeImage_GetLastError = LastError_ErrDesc
   
End Function

'=============================================================================================================
' FreeImage_SetLastError
'
' This function sets the error description that was previously returned from the "FI_OutputMessageProc"
' CALLBACK function to whatever you like.  This way, you can clear old error messages in preperation for new
' FreeImage calls.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' If an error occured which was returned via the CALLBACK funciton, this returns the error description
' If no error has occured, this returns a blank string
'=============================================================================================================
Public Sub FreeImage_SetLastError(ByVal strNewValue As String)
   
   LastError_ErrDesc = strNewValue
   
End Sub

'=============================================================================================================
' FreeImage_GrayscaleDIB
'
' This function takes a Device Independant Bitmap (DIB) returned from FreeImage and make a grayscale copy of it.
' This grayscale copy is then returned via the "Return_DIB" parameter.
'
' NOTE: The caller of this function is responsbile for cleaning up the returned DIB object (FreeImage_Unload)
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' DIB                           Specifies the Device Independant Bitmap (DIB) to convert to grayscale.  This
'                               is usually passed in as the result of a call to a FreeImage_* function.
' Return_DIB                    Returns a handle to a DIB image that is the same as the one passed in the "DIB"
'                               parameter (only in grayscale) which that can be used with FreeImage_* functions.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_GrayscaleDIB(ByVal DIB As Long, ByRef Return_DIB As Long) As Boolean
  
  Dim BmpInfo     As BITMAP
  Dim hDC_Temp    As Long
  Dim hBMP_Prev   As Long
  Dim hBITMAP     As Long
  Dim lngCounterX As Long
  Dim lngCounterY As Long
  Dim lngColorCur As Long
  Dim lngColorNew As Long
  
  ' Set the default values
  If Return_DIB <> 0 Then Call DeleteObject(Return_DIB): Return_DIB = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If DIB = 0 Then SetError "[FreeImage_GrayscaleDIB] No DIB specified to convert to grayscale": Exit Function
  
  ' Get a copy of the DIB in the form of a BITMAP to work with
  If FreeImage_BMP_From_DIB(DIB, hBITMAP) = False Then SetError "[FreeImage_GrayscaleDIB] FreeImage_BMP_From_DIB failed to convert the DIB to BITMAP": Exit Function
  
  ' Get the BITMAP's dimentions
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_GrayscaleDIB] GetObjectAPI failed to get the BITMAP's info reporting error code " & Err.LastDllError: GoTo ErrorOut
  
  ' Create a memory Device Context to work on
  If MemoryDC_Create(hDC_Temp) = False Then SetError "[FreeImage_GrayscaleDIB] Failed to create a DC to work with": GoTo ErrorOut
  
  ' Put the BITMAP into our temporary DC and save the previous BITMAP that was in there
  hBMP_Prev = SelectObject(hDC_Temp, hBITMAP)
  
  ' Loop through the pixels of the image and get the grayscale equivelants of each pixel
  For lngCounterY = 0 To BmpInfo.bmHeight - 1
    For lngCounterX = 0 To BmpInfo.bmWidth - 1
      
      ' Get the current pixel color
      lngColorCur = GetPixel(hDC_Temp, lngCounterX, lngCounterY)
      If lngColorCur <> -1 Then
        
        ' Get the grayscale equivelant of the pixel
        lngColorNew = 0.33 * (lngColorCur Mod 256) + _
                      0.59 * ((lngColorCur \ 256) Mod 256) + _
                      0.11 * ((lngColorCur \ 65536) Mod 256)
        lngColorNew = RGB(lngColorNew, lngColorNew, lngColorNew)
        
        ' Set the new grayscale pixel in the stead of the color pixel
        SetPixel hDC_Temp, lngCounterX, lngCounterY, lngColorNew
        
      End If
    Next lngCounterX
  Next lngCounterY
  
  ' Get the altered BITMAP back
  hBITMAP = SelectObject(hDC_Temp, hBMP_Prev)
  
  ' Clean up the DC we created
  DeleteDC hDC_Temp
  
  ' Convert the BITMAP to a DIB
  If FreeImage_DIB_From_BMP(hBITMAP, Return_DIB) = False Then SetError "[FreeImage_GrayscaleDIB] FreeImage_DIB_From_BMP failed to conver the grayscale bitmap to a DIB": Exit Function
  DeleteObject hBITMAP
  
  ' Function executed successfully
  FreeImage_GrayscaleDIB = True
  
  Exit Function
  
ErrorOut:
  
  If hBITMAP <> 0 Then DeleteObject hBITMAP
  If hDC_Temp <> 0 Then DeleteDC hDC_Temp
  
End Function

'=============================================================================================================
' FreeImage_GrayscaleBITMAP
'
' This function takes a handle to a BITMAP and makes a grayscale copy of it.  This grayscale copy is then
' returned via the "Return_BITMAP" parameter.
'
' NOTE: The caller of this function is responsbile for cleaning up the returned BITMAP GDI object (DeleteObject)
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the handle to a BITMAP to convert to grayscale.
' Return_BITMAP                 Returns a handle to a valid GDI BITMAP image that is the same as the one
'                               passed in "hBITMAP", only in grayscale.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_GrayscaleBITMAP(ByVal hBITMAP As Long, ByRef Return_BITMAP As Long) As Boolean

  Dim BmpInfo     As BITMAP
  Dim hDC_Temp    As Long
  Dim hBMP_Prev   As Long
  Dim hBITMAP_New As Long
  Dim lngCounterX As Long
  Dim lngCounterY As Long
  Dim lngColorCur As Long
  Dim lngColorNew As Long
  
  ' Set the default values
  If Return_BITMAP <> 0 Then Call DeleteObject(Return_BITMAP): Return_BITMAP = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then SetError "[FreeImage_GrayscaleBITMAP] No bitmap specified to convert to grayscale": Exit Function
  
  ' Get the BITMAP's dimentions
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_GrayscaleBITMAP] GetObjectAPI failed to get the BITMAP info reporting error code " & Err.LastDllError: Exit Function
  
  ' Get a copy of the original image to work with (and later return)
  hBITMAP_New = CopyImage(hBITMAP, IMAGE_BITMAP, BmpInfo.bmWidth, BmpInfo.bmHeight, 0)
  If hBITMAP_New = 0 Then SetError "[FreeImage_GrayscaleBITMAP] CopyImage failed to create a copy of the specified BITMAP with error code " & Err.LastDllError: Exit Function
  
  ' Create a memory Device Context to work on
  If MemoryDC_Create(hDC_Temp) = False Then SetError "[FreeImage_GrayscaleBITMAP] Failed to create a DC to work with": GoTo ErrorOut
  
  ' Put the BITMAP into our temporary DC and save the previous BITMAP that was in there
  hBMP_Prev = SelectObject(hDC_Temp, hBITMAP_New)
  
  ' Loop through the pixels of the image and get the grayscale equivelants of each pixel
  For lngCounterY = 0 To BmpInfo.bmHeight - 1
    For lngCounterX = 0 To BmpInfo.bmWidth - 1
      
      ' Get the current pixel color
      lngColorCur = GetPixel(hDC_Temp, lngCounterX, lngCounterY)
      If lngColorCur <> -1 Then
        
        ' Get the grayscale equivelant of the pixel
        lngColorNew = 0.33 * (lngColorCur Mod 256) + _
                      0.59 * ((lngColorCur \ 256) Mod 256) + _
                      0.11 * ((lngColorCur \ 65536) Mod 256)
        lngColorNew = RGB(lngColorNew, lngColorNew, lngColorNew)
        
        ' Set the new grayscale pixel in the stead of the color pixel
        SetPixel hDC_Temp, lngCounterX, lngCounterY, lngColorNew
        
      End If
    Next lngCounterX
  Next lngCounterY
  
  ' Get the altered BITMAP back
  hBITMAP_New = SelectObject(hDC_Temp, hBMP_Prev)
  
  ' Clean up the DC we created
  DeleteDC hDC_Temp
  
  ' Return the new bitmap
  Return_BITMAP = hBITMAP_New
  
  ' Function executed successfully
  FreeImage_GrayscaleBITMAP = True
  
  Exit Function
  
ErrorOut:
  
  If hBITMAP_New <> 0 Then DeleteObject hBITMAP_New
  If hDC_Temp <> 0 Then DeleteDC hDC_Temp
  
End Function

'=============================================================================================================
' FreeImage_BMP_From_DIB
'
' This function takes a Device Independant Bitmap (DIB) returned from FreeImage and converts it to a standard
' Win32 BITMAP.
'
' NOTE: The caller of this function is responsbile for cleaning up the returned BITMAP GDI object (DeleteObject)
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' DIB                           Specifies the Device Independant Bitmap (DIB) to convert to a BITMAP.  This
'                               is usually passed in as the result of a call to a FreeImage_* function.
' Return_hBITMAP                Returns a handle to a valid Win32 GDI BITMAP object.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_BMP_From_DIB(ByVal DIB As Long, ByRef Return_hBITMAP As Long) As Boolean
  
  Dim hDC_Screen As Long
  Dim lpBMI      As Long
  Dim lpBMIH     As Long
  Dim lpBits     As Long
  
  ' Clear return values
  If Return_hBITMAP <> 0 Then DeleteObject Return_hBITMAP: Return_hBITMAP = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If DIB = 0 Then SetError "[FreeImage_BMP_From_DIB] No DIB specified to convert": Exit Function
  
  ' Get the required information from the DIB
  lpBMIH = FreeImage_GetInfoHeader(DIB)
  If lpBMIH = 0 Then SetError "[FreeImage_BMP_From_DIB] FreeImage_GetInfoHeader failed to get the DIB header": Exit Function
  lpBMI = FreeImage_GetInfo(DIB)
  If lpBMI = 0 Then SetError "[FreeImage_BMP_From_DIB] FreeImage_GetInfo failed to get the DIB info": Exit Function
  lpBits = FreeImage_GetBits(DIB)
  If lpBits = 0 Then SetError "[FreeImage_BMP_From_DIB] FreeImage_GetBits failed to get the DIB's bits": Exit Function
  
  ' Get a reference to the current display's Device Context (DC)
  hDC_Screen = GetDC(GetDesktopWindow)
  If hDC_Screen = 0 Then SetError "[FreeImage_BMP_From_DIB] GetDC failed to get the DC to work with reporting error code " & Err.LastDllError: Exit Function
  
  ' Create a new BITMAP from the specified DIB
  Return_hBITMAP = CreateDIBitmap(hDC_Screen, ByVal lpBMIH, CBM_INIT, ByVal lpBits, ByVal lpBMI, DIB_RGB_COLORS)
  If Return_hBITMAP <> 0 Then
    FreeImage_BMP_From_DIB = True
  Else
    SetError "[FreeImage_BMP_From_DIB] CreateDIBitmap failed to create the BITMAP from the DIB reporting error code " & Err.LastDllError
  End If
  
CleanUp:
  ReleaseDC GetDesktopWindow, hDC_Screen
  
End Function

'=============================================================================================================
' FreeImage_DIB_From_BMP
'
' This function takes a standard Win32 BITMAP and converts it to a Device Independant Bitmap (DIB) that can
' be used with FreeImage.
'
' NOTE: The caller of this function is responsbile for cleaning up the returned BITMAP GDI object (DeleteObject)
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the handle to a valid Win32 GDI BITMAP object.
' Return_DIB                    Returns the handle to the Device Independant Bitmap (DIB) that was created
'                               from the specified BITMAP.  This can be used with FreeImage_* functions.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_DIB_From_BMP(ByVal hBITMAP As Long, ByRef Return_DIB As Long) As Boolean
  
  Dim BmpInfo    As BITMAP
  Dim hDC_Screen As Long
  Dim TempDIB    As Long
  Dim lpBits     As Long
  Dim lpBMI      As Long
  Dim TheWidth   As Long
  Dim TheHeight  As Long
  Dim TheBPP     As Integer
  
  ' Clear return values
  If Return_DIB <> 0 Then FreeImage_Unload Return_DIB: Return_DIB = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then SetError "[FreeImage_DIB_From_BMP] No bitmap specified to convert": Exit Function
  
  ' Validate the BITMAP and get the information we need from it
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_DIB_From_BMP] GetObjectAPI failed with error code " & Err.LastDllError: Exit Function
  TheWidth = BmpInfo.bmWidth
  TheHeight = BmpInfo.bmHeight
 'TheBPP = BmpInfo.bmPlanes * BmpInfo.bmBitsPixel
  TheBPP = 32 ' This is hard coded as a "work around" to a problem in getting the correct color depth of images returned by FreeImage
  
  ' Get a reference to the current display's Device Context (DC)
  hDC_Screen = GetDC(GetDesktopWindow)
  If hDC_Screen = 0 Then SetError "[FreeImage_DIB_From_BMP] GetDC failed to get a DC to work with reporting error code " & Err.LastDllError: Exit Function
  
  ' Allocate a DIB and get the information from it we need
  TempDIB = FreeImage_Allocate(TheWidth, TheHeight, TheBPP)
  If TempDIB = 0 Then SetError "[FreeImage_DIB_From_BMP] FreeImage_Allocate failed to create a DIB to work with": GoTo CleanUp
  lpBMI = FreeImage_GetInfo(TempDIB)
  If lpBMI = 0 Then Call FreeImage_Free(TempDIB): SetError "[FreeImage_DIB_From_BMP] FreeImage_GetInfo failed to get the DIB info": GoTo CleanUp
  lpBits = FreeImage_GetBits(TempDIB)
  If lpBits = 0 Then Call FreeImage_Free(TempDIB): SetError "[FreeImage_DIB_From_BMP] FreeImage_GetBits failed to get the DIB bits": GoTo CleanUp
  
  ' Convert the DIB to a BITMAP
  If GetDIBits(hDC_Screen, hBITMAP, 0, TheHeight, ByVal lpBits, ByVal lpBMI, DIB_RGB_COLORS) <> 0 Then
    FreeImage_DIB_From_BMP = True
    Return_DIB = TempDIB
  Else
    SetError "[FreeImage_DIB_From_BMP] GetDIBits failed to get the DIB's bits reporting error code " & Err.LastDllError
    FreeImage_Free TempDIB
  End If
  
CleanUp:
  ReleaseDC GetDesktopWindow, hDC_Screen
  
End Function

'=============================================================================================================
' FreeImage_RenderDIB
'
' This function takes a Device Indpendant Bitmap (DIB) returned from FreeImage and renders it to the specified
' Device Context (DC).
'
' NOTE: This function can be used to render a DIB onto a memory Device Context (DC) or a window-based DC
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' DIB                           Specifies the Device Independant Bitmap to render into the specified DC
' Dest_hDC                      Specifies the memory based DC (Device Context) or window based DC to render onto
' Dest_X                        Optioanl. Specifies the X coordinate (Left) of where the image should be rendered
' Dest_Y                        Optioanl. Specifies the Y coordinate (Top) of where the image should be rendered
' Src_Width                     Optioanl. Specifies the width of the DIB to render.  If this isn't specified,
'                               this function attempts to get the width from the image passed in.
' Src_Height                    Optioanl. Specifies the height of the DIB to render.  If this isn't specified,
'                               this function attempts to get the height from the image passed in.
' Return_ErrNum                 Optioanl. If an error occurs, this returns the Win32 error number.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_RenderDIB(ByVal DIB As Long, ByVal Dest_hDC As Long, Optional ByVal Dest_X As Long = 0, Optional ByVal Dest_Y As Long = 0, Optional ByVal Src_Width As Long = -1, Optional ByVal Src_Height As Long = -1) As Boolean
  
  Dim lpBMI     As Long
  Dim lpBits    As Long
  Dim TheWidth  As Long
  Dim TheHeight As Long
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If DIB = 0 Then
    SetError "[FreeImage_RenderDIB] No DIB specified to render": Exit Function
  ElseIf GetObjectType(Dest_hDC) <> OBJ_DC And GetObjectType(Dest_hDC) <> OBJ_MEMDC Then
    SetError "[FreeImage_RenderDIB] DC specified was invalid": Exit Function
  End If
  
  ' Get the width to use
  If Src_Width > 0 Then
    TheWidth = Src_Width
  Else
    TheWidth = FreeImage_GetWidth(DIB)
    If TheWidth < 1 Then SetError "[FreeImage_RenderDIB] Could not get the bitmap's width": Exit Function
  End If
  
  ' Get the height to use
  If Src_Height > 0 Then
    TheHeight = Src_Height
  Else
    TheHeight = FreeImage_GetHeight(DIB)
    If TheHeight < 1 Then SetError "[FreeImage_RenderDIB] Could not get the bitmap's height": Exit Function
  End If
  
  ' Get the bits of the DIB
  lpBits = FreeImage_GetBits(DIB)
  If lpBits = 0 Then SetError "[FreeImage_RenderDIB] Could not get the DIB's bits": Exit Function
  
  ' Get the BITMAP info for the DIB
  lpBMI = FreeImage_GetInfo(DIB)
  If lpBMI = 0 Then SetError "[FreeImage_RenderDIB] Could not get the DIB's information": Exit Function
  
  ' Render the DIB to the specified device
  If SetDIBitsToDevice(Dest_hDC, Dest_X, Dest_Y, TheWidth, TheHeight, 0, 0, 0, TheHeight, ByVal lpBits, ByVal lpBMI, DIB_RGB_COLORS) = 0 Then
    SetError "[FreeImage_RenderDIB] SetDIBitsToDevice failed to render the DIB to the specified DC with error code " & Err.LastDllError
  Else
    FreeImage_RenderDIB = True
  End If
  
End Function

'=============================================================================================================
' FreeImage_RenderBitmap
'
' This function takes a standard Win32 BITMAP and renders it to the specified Device Context (DC).
'
' NOTE: This function can be used to render a BITMAP onto a memory Device Context (DC) or a window-based DC
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the Win32 GDI BITMAP object to render into the specified DC
' Dest_hDC                      Specifies the memory based DC (Device Context) or window based DC to render onto
' Dest_X                        Optioanl. Specifies the X coordinate (Left) of where the image should be rendered
' Dest_Y                        Optioanl. Specifies the Y coordinate (Top) of where the image should be rendered
' Src_Width                     Optioanl. Specifies the width of the BITMAP to render.  If this isn't specified,
'                               this function attempts to get the width from the image passed in.
' Src_Height                    Optioanl. Specifies the height of the BITMAP to render.  If this isn't specified,
'                               this function attempts to get the height from the image passed in.
' blnInvertBitmapColors         Optional. If set to TRUE, the colors of the specified image will be inverted.
' Return_ErrNum                 Optioanl. If an error occurs, this returns the Win32 error number.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_RenderBitmap(ByVal hBITMAP As Long, ByVal Dest_hDC As Long, Optional ByVal Dest_X As Long = 0, Optional ByVal Dest_Y As Long = 0, Optional ByVal Src_Width As Long = -1, Optional ByVal Src_Height As Long = -1, Optional ByVal blnInvertBitmapColors As Boolean = False) As Boolean
  
  Dim hdc       As Long
  Dim hBMP_Prev As Long
  Dim TheWidth  As Long
  Dim TheHeight As Long
  Dim TheFlags  As Long
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then
    SetError "[FreeImage_RenderBitmap] No bitmap specified to render": Exit Function
  ElseIf GetObjectType(hBITMAP) <> OBJ_BITMAP Then
    SetError "[FreeImage_RenderBitmap] Handle specified is not a valid BITMAP": Exit Function
  ElseIf GetObjectType(Dest_hDC) <> OBJ_DC And GetObjectType(Dest_hDC) <> OBJ_MEMDC Then
    SetError "[FreeImage_RenderBitmap] DC specified is not valid": Exit Function
  End If
  
  ' Get the width to use
  If Src_Width > 0 Then
    TheWidth = Src_Width
  Else
    TheWidth = FreeImage_GetBitmapWidth(hBITMAP)
    If TheWidth < 1 Then SetError "[FreeImage_RenderBitmap] Could not get bitmap's width": Exit Function
  End If
  
  ' Get the height to use
  If Src_Height > 0 Then
    TheHeight = Src_Height
  Else
    TheHeight = FreeImage_GetBitmapHeight(hBITMAP)
    If TheHeight < 1 Then SetError "[FreeImage_RenderBitmap] Could not get bitmap's height": Exit Function
  End If
  
  ' Create a DC to draw with
  If MemoryDC_Create(hdc) = False Then SetError "[FreeImage_RenderBitmap] Coult not create a DC to work with": Exit Function
  
  ' Put the BITMAP into the newly created memory DC
  hBMP_Prev = SelectObject(hdc, hBITMAP)
  
  ' Draw onto the specified DC
  TheFlags = vbSrcCopy
  If blnInvertBitmapColors = True Then TheFlags = vbNotSrcCopy
  If BitBlt(Dest_hDC, Dest_X, Dest_Y, TheWidth, TheHeight, hdc, 0, 0, TheFlags) = FALSE_ Then
    SetError "[FreeImage_RenderBitmap] BitBlt failed to render the BITMAP to the DC with error code " & Err.LastDllError
  Else
    FreeImage_RenderBitmap = True
  End If
  
  ' Take our bitmap out of the DC
  SelectObject hdc, hBMP_Prev
  
  ' Clean up the tempoary memory DC
  MemoryDC_Destroy hdc
  
End Function

'=============================================================================================================
' FreeImage_GetBitmapWidth
'
' This function returns the width of the specified Win32 BITMAP in pixels.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the Win32 GDI BITMAP object to get the information from
'
' Return:
' ¯¯¯¯¯¯¯
' If successful, returns the width in pixels of the specified BITMAP
' If failed, returns ZERIO (0)
'=============================================================================================================
Public Function FreeImage_GetBitmapWidth(ByVal hBITMAP As Long) As Long
  
  Dim BmpInfo As BITMAP
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then SetError "[FreeImage_GetBitmapWidth] No bitmap specified to get the info for": Exit Function
  If GetObjectType(hBITMAP) <> OBJ_BITMAP Then SetError "[FreeImage_GetBitmapWidth] Handle specified is not a valid BITMAP": Exit Function
  
  ' Get the information for the bitmap
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_GetBitmapWidth] GetObjectAPI failed with error code " & Err.LastDllError: Exit Function
  FreeImage_GetBitmapWidth = BmpInfo.bmWidth
  
End Function

'=============================================================================================================
' FreeImage_GetBitmapHeight
'
' This function returns the height of the specified Win32 BITMAP in pixels.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the Win32 GDI BITMAP object to get the information from
'
' Return:
' ¯¯¯¯¯¯¯
' If successful, returns the width in pixels of the specified BITMAP
' If failed, returns ZERIO (0)
'=============================================================================================================
Public Function FreeImage_GetBitmapHeight(ByVal hBITMAP As Long) As Long
  
  Dim BmpInfo As BITMAP
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then SetError "[FreeImage_GetBitmapHeight] No bitmap specified to get the info for": Exit Function
  If GetObjectType(hBITMAP) <> OBJ_BITMAP Then SetError "[FreeImage_GetBitmapHeight] Handle specified is not a valid BITMAP": Exit Function
  
  ' Get the information for the bitmap
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_GetBitmapHeight] GetObjectAPI failed with error code " & Err.LastDllError: Exit Function
  FreeImage_GetBitmapHeight = BmpInfo.bmHeight
  
End Function

'=============================================================================================================
' FreeImage_GetBitmapInfo
'
' This function returns the information of the specified Win32 BITMAP.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' hBITMAP                       Specifies the valid Win32 BITMAP object to get information about.
' Return_Width                  Optional. Returns the width of the specified BITMAP in pixels.
' Return_Height                 Optional. Returns the height of the specified BITMAP in pixels.
' Return_Type                   Optional. Returns the type of BITMAP specified.
' Return_BitsPerPixel           Optional. Returns the color depth (bits per pixel) of the specified BITMAP.
'                               1  = Black and white
'                               4  = 16 Colors
'                               8  = 256 Colors
'                               16 = 16 bit Color (65536 Colors)
'                               24 = 24 bit Color - True Color (16777216 Colors)
'                               32 = 32 bit Color - High Color (4294967296 Colors)
' Return_Planes                 Optional. Specifies the count of color planes.  This is usually 1.
' Return_PointerToBits          Optional. Returns a pointer to the memory location of the first bit of the
'                               bit array that makes up the colors of the specified BITMAP.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_GetBitmapInfo(ByVal hBITMAP As Long, Optional ByRef Return_Width As Long, Optional ByRef Return_Height As Long, Optional ByRef Return_Type As Long, Optional ByRef Return_BitsPerPixel As Integer, Optional ByRef Return_Planes As Integer, Optional ByRef Return_PointerToBits As Long) As Boolean
  
  Dim BmpInfo As BITMAP
  
  ' Clear return values
  Return_BitsPerPixel = -1
  Return_Height = -1
  Return_Planes = -1
  Return_PointerToBits = -1
  Return_Type = -1
  Return_Width = -1
  LastError_ErrDesc = ""
  
  ' Validate parameters
  If hBITMAP = 0 Then SetError "[FreeImage_GetBitmapInfo] No bitmap specified to get the info for": Exit Function
  If GetObjectType(hBITMAP) <> OBJ_BITMAP Then SetError "[FreeImage_GetBitmapInfo] Handle specified is not a valid bitmap": Exit Function
  
  ' Get the information for the bitmap
  If GetObjectAPI(hBITMAP, Len(BmpInfo), BmpInfo) = 0 Then SetError "[FreeImage_GetBitmapInfo] GetObjectAPI failed with error code " & Err.LastDllError: Exit Function
  
  ' Return the information
  With BmpInfo
    Return_Height = .bmHeight
    Return_Planes = .bmPlanes
    Return_Type = .bmType
    Return_Width = .bmWidth
    Return_BitsPerPixel = .bmPlanes * .bmBitsPixel
    Return_PointerToBits = .bmBits
  End With
  FreeImage_GetBitmapInfo = True
  
End Function

'=============================================================================================================
' FreeImage_MultiPage_Load
'
' This function takes the specified file and returns back an array of BITMAP handles representing the
' BITMAP(s) that make it up
'
' NOTE: The caller is responsible for destroying the returned BITMAP(s) by calling the Win32 DeleteObject API
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName                      File name or file path to the Multi-Page image file you wish to open
' Return_HBITMAPs               Returns an array of handles to the BITMAPS(s) that make up the original
'                               Multi-Page image.
' Return_DibCount               Returns the number of images found within the specified image file
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_Load(ByVal FileName As String, ByRef Return_HBITMAPs() As Long, ByRef Return_BitmapCount As Long) As Boolean
  
  Dim ImageFormat As FREE_IMAGE_FORMAT
  Dim strTempPath As String
  Dim MBMP        As Long
  Dim Page        As Long
  Dim PageCount   As Long
  Dim hDibTemp    As Long
  Dim hBitmapTemp As Long
  Dim lngCounter  As Long
  
  ' Clear return values
  Erase Return_HBITMAPs
  Return_BitmapCount = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName = Trim$(FileName)
  If FileName = "" Then SetError "[FreeImage_MultiPage_Load] No file specified to load": Exit Function
  If Dir(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then SetError "[FreeImage_MultiPage_Load] File not found": Exit Function
  ImageFormat = FreeImage_GetFileType(FileName)
  If ImageFormat = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_Load] Unable to determine the file type of the specified multipage bitmap": Exit Function
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempPath = StrConv(FileName, vbFromUnicode)
  
  ' Open the file as a Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat, StrPtr(strTempPath), FALSE_, TRUE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_Load] FreeImage_OpenMultiBitmap failed to open the file": Exit Function
  
  ' Get the total number of pages in the image
  PageCount = FreeImage_GetPageCount(MBMP)
  If PageCount < 1 Then SetError "[FreeImage_MultiPage_Load] FreeImage_GetPageCount failed to get the number of images in the file": GoTo CleanUp
  
  ' Loop through the pages and transfer the DIB pages returned into BITMAPS and return them as an array of HANDLE(s)
  For lngCounter = 0 To PageCount - 1
    
    ' Lock the current page to work with it
    Page = FreeImage_LockPage(MBMP, lngCounter)
    
    ' Make sure the specified page is valid
    If Page <> 0 Then
      
      ' Make a copy of the page and return it in the array
      hDibTemp = FreeImage_Clone(Page)
      
      ' Convert the current DIB image to a BITMAP
      If FreeImage_BMP_From_DIB(hDibTemp, hBitmapTemp) = True Then
        If hBitmapTemp <> 0 Then
          ReDim Preserve Return_HBITMAPs(0 To Return_BitmapCount) As Long
          Return_HBITMAPs(Return_BitmapCount) = hBitmapTemp
          Return_BitmapCount = Return_BitmapCount + 1
        End If
      End If
      
      ' Clean up the temporary DIB
      FreeImage_Unload hDibTemp
      hDibTemp = 0
      hBitmapTemp = 0
      
      ' Unlock the current page
      FreeImage_UnlockPage MBMP, Page, FALSE_
    End If
    
  Next
  
  FreeImage_MultiPage_Load = True
  
CleanUp:
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function

'=============================================================================================================
' FreeImage_MultiPage_LoadEx
'
' This function takes the specified file and returns back an array of DIB handles representing the Device
' Independant BITMAP(s) that make it up.
'
' NOTE: The caller is responsible for destroying the returned DIB(s) by calling the "FreeImage_Free" or
' "FreeImage_Unload" API
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName                      File name or file path to the Multi-Page image file you wish to open
' Return_DIBs                   Returns an array of handles to the Device Independant Bitamp(s) that make up
'                               the original Multi-Page image.
' Return_DibCount               Returns the number of images found within the specified image file
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_LoadEx(ByVal FileName As String, ByRef Return_DIBs() As Long, ByRef Return_DibCount As Long) As Boolean
  
  Dim ImageFormat As FREE_IMAGE_FORMAT
  Dim strTempPath As String
  Dim MBMP        As Long
  Dim Page        As Long
  Dim PageCount   As Long
  Dim hDibTemp    As Long
  Dim lngCounter  As Long
  
  ' Clear return values
  Erase Return_DIBs
  Return_DibCount = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName = Trim$(FileName)
  If FileName = "" Then SetError "[FreeImage_MultiPage_LoadEx] No file specified to load": Exit Function
  If Dir(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_LoadEx] File Not Found": Exit Function
  ImageFormat = FreeImage_GetFileType(FileName)
  If ImageFormat = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_LoadEx] Unable to determine the file type of the specified multipage bitmap": Exit Function
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempPath = StrConv(FileName, vbFromUnicode)
  
  ' Open the file as a Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat, StrPtr(strTempPath), FALSE_, TRUE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_LoadEx] FreeImage_OpenMultiBitmap failed to open the file": Exit Function
  
  ' Get the total number of pages in the image
  PageCount = FreeImage_GetPageCount(MBMP)
  If PageCount < 1 Then SetError "[FreeImage_MultiPage_LoadEx] FreeImage_GetPageCount failed to get the page count": GoTo CleanUp
  
  ' Loop through the pages and transfer the DIB pages returned into BITMAPS and return them as an array of HANDLE(s)
  For lngCounter = 0 To PageCount - 1
    
    ' Lock the current page to work with it
    Page = FreeImage_LockPage(MBMP, lngCounter)
    
    ' Make sure the specified page is valid
    If Page <> 0 Then
      
      ' Make a copy of the page and return it in the array
      hDibTemp = FreeImage_Clone(Page)
      If hDibTemp <> 0 Then
        ReDim Preserve Return_DIBs(0 To Return_DibCount) As Long
        Return_DIBs(Return_DibCount) = hDibTemp
        Return_DibCount = Return_DibCount + 1
      End If
      
      ' Unlock the current page
      FreeImage_UnlockPage MBMP, Page, FALSE_
    End If
    
  Next
  
  FreeImage_MultiPage_LoadEx = True
  
CleanUp:
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function


'=============================================================================================================
' FreeImage_MultiPage_PageCount
'
' This function takes the specified multipage file and returns how many "pages" or images are in it.
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName_Multipage            File name or file path to the Multi-Page image
' Return_PageCount              Returns the number of pages within the specified Multi-Page image
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_PageCount(ByVal FileName_Multipage As String, ByRef Return_PageCount As Long) As Boolean
  
  Dim ImageFormat_M  As FREE_IMAGE_FORMAT
  Dim strTempMulti   As String
  Dim MBMP           As Long
  
  ' Clear return values
  Return_PageCount = 0
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName_Multipage = Trim$(FileName_Multipage)
  If FileName_Multipage = "" Then SetError "[FreeImage_MultiPage_PageCount] No multipage bitmap file specified": Exit Function
  If Dir(FileName_Multipage, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_PageCount] Multipage Bitmap File Not Found": Exit Function
  ImageFormat_M = FreeImage_GetFileType(FileName_Multipage)
  If ImageFormat_M = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_PageCount] Unable to determine the file type of the specified multipage bitmap": Exit Function
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempMulti = StrConv(FileName_Multipage, vbFromUnicode)
  
  ' Open the Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat_M, StrPtr(strTempMulti), FALSE_, TRUE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_PageCount] FreeImage_OpenMultiBitmap failed to open the multipage file": Exit Function
  
  ' Get the total number of pages in the multipage image
  Return_PageCount = FreeImage_GetPageCount(MBMP)
  
CleanUp:
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function


'=============================================================================================================
' FreeImage_MultiPage_Add
'
' This function takes the specified multipage file and adds the specified image to it.
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName_Multipage            File name or file path to the Multi-Page image file you wish to add to.
' DIB_New                       Optional. Specifies the DIB to add to the specified Multi-Page image file.
'                               If this is not specified, the "FileName_New" must be.
' FileName_New                  Optional. File name or file path to the image you wish to add to the multi-
'                               page bitmap.  If this is not specified, the "DIB_New" parameter must be.
' NewFileOpenFlags              Optional. Flags to use when opening the "FileName_New" image.  This is passed
'                               to the "Flags" parameter of the "FreeImage_Load1" function.
' PageNumber                    Optional. If set to anything greater than negative one (-1), the specified
'                               image will be inserted at the specified index.  If left at -1, the specified
'                               image will be inserted at the end of the Multi-Page image.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_Add(ByVal FileName_Multipage As String, Optional ByVal DIB_New As Long, Optional ByVal FileName_New As String, Optional ByVal NewFileOpenFlags As Long, Optional ByVal PageNumber As Long = -1) As Boolean
  
  Dim ImageFormat_M  As FREE_IMAGE_FORMAT
  Dim ImageFormat_A  As FREE_IMAGE_FORMAT
  Dim strTempMulti   As String
  Dim MBMP           As Long
  Dim DIB            As Long
  Dim ImgCountBefore As Long
  Dim ImgCountAfter  As Long
  Dim blnCleanUpDIB  As Boolean
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName_Multipage = Trim$(FileName_Multipage)
  FileName_New = Trim$(FileName_New)
  If FileName_Multipage = "" Then SetError "[FreeImage_MultiPage_Add] No multipage bitmap file specified to add to": Exit Function
  If Dir(FileName_Multipage, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_Add] Multipage Bitmap File Not Found": Exit Function
  ImageFormat_M = FreeImage_GetFileType(FileName_Multipage)
  If ImageFormat_M = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_Add] Unable to determine the file type of the specified multipage bitmap": Exit Function
  If DIB_New = 0 And FileName_New = "" Then SetError "[FreeImage_MultiPage_Add] No multipage bitmap file specified to add to": Exit Function
  
  ' If the DIB is specified, use it
  If DIB_New <> 0 Then
    DIB = DIB_New
    blnCleanUpDIB = False
    
  ' If the file name is specified, open it and use it
  Else
    If FileName_New = "" Then SetError "[FreeImage_MultiPage_Add] No image file specified to add to multipage": Exit Function
    If Dir(FileName_New, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_Add] File Not Found": Exit Function
    ImageFormat_A = FreeImage_GetFileType(FileName_New)
    If ImageFormat_A = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_Add] Unable to determine the file type of the file specified to add to the multipage bitmap": Exit Function
    
    ' Open the new image
    DIB = FreeImage_Load(ImageFormat_A, FileName_New, NewFileOpenFlags)
    If DIB = 0 Then SetError "[FreeImage_MultiPage_Add] FreeImage_Load failed to open the new file": Exit Function
    blnCleanUpDIB = True
  End If
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempMulti = StrConv(FileName_Multipage, vbFromUnicode)
  
  ' Open the Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat_M, StrPtr(strTempMulti), FALSE_, FALSE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_Add] FreeImage_OpenMultiBitmap failed to open the multipage file": Exit Function
  
  ' Get the total number of pages in the multipage image BEFORE we add one
  ImgCountBefore = FreeImage_GetPageCount(MBMP)
  
  ' If "PageNumber" is negative one, append the specified DIB to the end of the MultiPage
  If PageNumber = -1 Then
    Call FreeImage_AppendPage(MBMP, DIB)
    
  ' Add the DIB at the specified index
  ElseIf PageNumber > -1 Then
    If PageNumber >= (ImgCountBefore - 1) Then
      Call FreeImage_AppendPage(MBMP, DIB)
    Else
      Call FreeImage_InsertPage(MBMP, PageNumber, DIB)
    End If
    
  ' Invalid PageNumber specified
  ElseIf PageNumber < -1 Then
    SetError "[FreeImage_MultiPage_Add] Invalid page number specified"
  End If
  
  ' Check for errors
  If LastError_ErrDesc = "" Then
    
    ' Get the total number of pages in the multipage image AFTER we add one.  Should be one more than before.
    ImgCountAfter = FreeImage_GetPageCount(MBMP)
    If ImgCountAfter > ImgCountBefore Then
      FreeImage_MultiPage_Add = True
    Else
       SetError "[FreeImage_MultiPage_Add] Failed to correctly add the specified image"
    End If
  End If
  
CleanUp:
  
  ' Clean up the DIB we created
  If DIB <> 0 And blnCleanUpDIB = True Then FreeImage_Unload DIB
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function


'=============================================================================================================
' FreeImage_MultiPage_Move
'
' This function takes the specified multipage file and moves the specified image from one location within
' it to another.
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName_Multipage            File name or file path to the Multi-Page image file you wish to add to.
' PageNumber_Source             Specifies the index of the page to move.
' PageNumber_Destination        Specifies the index of the page to move the source to.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_Move(ByVal FileName_Multipage As String, ByVal PageNumber_Source As Long, ByVal PageNumber_Destination As Long) As Boolean
  
  Dim ImageFormat_M  As FREE_IMAGE_FORMAT
  Dim strTempMulti   As String
  Dim MBMP           As Long
  Dim PageCount      As Long
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName_Multipage = Trim$(FileName_Multipage)
  If FileName_Multipage = "" Then SetError "[FreeImage_MultiPage_Move] No multipage bitmap file specified": Exit Function
  If Dir(FileName_Multipage, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_Move] Multipage Bitmap File Not Found": Exit Function
  ImageFormat_M = FreeImage_GetFileType(FileName_Multipage)
  If ImageFormat_M = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_Move] Unable to determine the file type of the specified multipage bitmap": Exit Function
  
  ' If the source and destination indexes are the same, exit without error
  If PageNumber_Source = PageNumber_Destination Then
    FreeImage_MultiPage_Move = True
    Exit Function
  End If
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempMulti = StrConv(FileName_Multipage, vbFromUnicode)
  
  ' Open the Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat_M, StrPtr(strTempMulti), FALSE_, FALSE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_Move] FreeImage_OpenMultiBitmap failed to open the multipage file": Exit Function
  
  ' Get the total number of pages in the multipage image
  PageCount = FreeImage_GetPageCount(MBMP)
  
  ' Check if the specified source index is valid
  If PageNumber_Source < 0 Or PageNumber_Source > (PageCount - 1) Then
    SetError "[FreeImage_MultiPage_Move] Invalid source page number specified"
    
  ' Check if the specified destination index is valid
  ElseIf PageNumber_Destination < 0 Or PageNumber_Destination > (PageCount - 1) Then
    SetError "[FreeImage_MultiPage_Move] Invalid destination page number specified"
    
  ' Move the specified page
  Else
    If FreeImage_MovePage(MBMP, PageNumber_Source, PageNumber_Destination) = FALSE_ Then
      SetError "[FreeImage_MultiPage_Move] FreeImage_MovePage failed to move the specified page"
    Else
      FreeImage_MultiPage_Move = True
    End If
  End If
  
CleanUp:
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function


'=============================================================================================================
' FreeImage_MultiPage_Delete
'
' This function takes the specified multipage file and deletes the specified image from it.
'
' NOTE: As of 09/04/02, TIFF images are the only "MultiPage" format that is supported.
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' FileName_Multipage            File name or file path to the Multi-Page image file you wish to add to.
' PageNumber                    Specifies the index of the page to delete.
'
' Return:
' ¯¯¯¯¯¯¯
' Returns TRUE if the function executed successfully.
' Returns FALSE if the function failed to execute successfully.
'=============================================================================================================
Public Function FreeImage_MultiPage_Delete(ByVal FileName_Multipage As String, ByVal PageNumber As Long) As Boolean
  
  Dim ImageFormat_M  As FREE_IMAGE_FORMAT
  Dim strTempMulti   As String
  Dim MBMP           As Long
  Dim ImgCountBefore As Long
  Dim ImgCountAfter  As Long
  
  ' Clear return values
  LastError_ErrDesc = ""
  
  ' Validate parameters
  FileName_Multipage = Trim$(FileName_Multipage)
  If FileName_Multipage = "" Then SetError "[FreeImage_MultiPage_Delete] No multipage bitmap file specified to add to": Exit Function
  If Dir(FileName_Multipage, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then LastError_ErrDesc = "[FreeImage_MultiPage_Delete] Multipage Bitmap File Not Found": Exit Function
  ImageFormat_M = FreeImage_GetFileType(FileName_Multipage)
  If ImageFormat_M = FIF_UNKNOWN Then SetError "[FreeImage_MultiPage_Delete] Unable to determine the file type of the specified multipage bitmap": Exit Function
  
  ' Convert the file path to a static string we can use with the API call.  Whenever working with multipage bitmaps, this must be done.
  strTempMulti = StrConv(FileName_Multipage, vbFromUnicode)
  
  ' Open the Multi-Page
  MBMP = FreeImage_OpenMultiBitmap(ImageFormat_M, StrPtr(strTempMulti), FALSE_, FALSE_, FALSE_)
  If MBMP = 0 Then SetError "[FreeImage_MultiPage_Delete] FreeImage_OpenMultiBitmap failed to open the multipage file": Exit Function
  
  ' Get the total number of pages in the multipage image BEFORE we delete one
  ImgCountBefore = FreeImage_GetPageCount(MBMP)
  
  ' If "PageNumber" is a negative value, or is greater than the total number of pages, throw error
  If PageNumber < 0 Or PageNumber > (ImgCountBefore - 1) Then
    SetError "[FreeImage_MultiPage_Delete] Invalid page number specified"
    
  ' Delete the specified page
  ElseIf PageNumber > -1 Then
    Call FreeImage_DeletePage(MBMP, PageNumber)
  End If
  
  ' Check for errors
  If LastError_ErrDesc = "" Then
    
    ' Get the total number of pages in the multipage image AFTER we delete one.  Should be one less than before.
    ImgCountAfter = FreeImage_GetPageCount(MBMP)
    If ImgCountAfter < ImgCountBefore Then
      FreeImage_MultiPage_Delete = True
    Else
       SetError "[FreeImage_MultiPage_Delete] Failed to correctly delete the specified page"
    End If
  End If
  
CleanUp:
  
  ' Close the multibitmap
  If MBMP <> 0 Then FreeImage_CloseMultiBitmap MBMP
  
End Function


'=============================================================================================================
' BITMAPINFO_From_Ptr
'
' This function takes the specified pointer to a BITMAPINFO structure and returns that structure
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' lngPointer                    A long value that represents a pointer to the BITMAPINFO information to get
' Return_BITMAPINFO             Returns the BITMAPINFO information (if the pointer is a valid pointer to it)
'
' Return:
' ¯¯¯¯¯¯¯
' Nothing
'=============================================================================================================
Public Sub BITMAPINFO_From_Ptr(ByVal lngPointer As Long, ByRef Return_BITMAPINFO As BITMAPINFO)
  
  ' Validate parameters
  If lngPointer = 0 Then Exit Sub
  
  ' Copy the memory from the pointer to the structure
  CopyMemory Return_BITMAPINFO, ByVal lngPointer, Len(Return_BITMAPINFO)
  
End Sub

'=============================================================================================================
' BITMAPINFOHEADER_From_Ptr
'
' This function takes the specified pointer to a BITMAPINFOHEADER structure and returns that structure
'
' Parameter:                    Use:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' lngPointer                    A long value that represents a pointer to the BITMAPINFOHEADER information to get
' Return_BITMAPINFOHEADER       Returns the BITMAPINFOHEADER information (if the pointer is a valid pointer to it)
'
' Return:
' ¯¯¯¯¯¯¯
' Nothing
'=============================================================================================================
Public Sub BITMAPINFOHEADER_From_Ptr(ByVal lngPointer As Long, ByRef Return_BITMAPINFOHEADER As BITMAPINFOHEADER)
  
  ' Validate parameters
  If lngPointer = 0 Then Exit Sub
  
  ' Copy the memory from the pointer to the structure
  CopyMemory Return_BITMAPINFOHEADER, ByVal lngPointer, Len(Return_BITMAPINFOHEADER)
  
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXX            HELPER  FUNCTION  DECLARATIONS  (USED LOCALLY ONLY)          XXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' This function creates a memory Device Context (DC) compatible with the user's current display.  This DC can be used to perform BITMAP opperations.
Private Function MemoryDC_Create(ByRef Return_hDC As Long) As Boolean
  
  Dim hDC_Screen  As Long
  
  ' Clear return value
  Return_hDC = 0
  
  ' Get a reference to the current display's DC
  hDC_Screen = GetDC(GetDesktopWindow)
  If hDC_Screen = 0 Then Exit Function
  
  ' Create a DC that is compatible with the current display
  Return_hDC = CreateCompatibleDC(hDC_Screen)
  If Return_hDC <> 0 Then MemoryDC_Create = True
  
  ' Release the reference to the display's DC
  ReleaseDC GetDesktopWindow, hDC_Screen
  
End Function

' This function takes the specified memory Device Context (DC) and deletes it and any bitmaps that are currently selected into it.
Private Function MemoryDC_Destroy(ByRef hdc As Long, Optional ByRef hBITMAP_Previous As Long) As Boolean
  
  ' If there's a previous BITMAP, put it back before destroying the DC
  If hBITMAP_Previous <> 0 Then
    DeleteObject SelectObject(hdc, hBITMAP_Previous)
    hBITMAP_Previous = 0
  End If
  
  ' Delete the DC and remove the reference
  If DeleteDC(hdc) <> FALSE_ Then
    hdc = 0
    MemoryDC_Destroy = True
  End If
  
End Function

' This routine makes it easy to set the last error, but not overwrite any previous error messages
Private Sub SetError(ByVal strErrDesc As String, Optional ByVal blnOnlySetIfBlank As Boolean = True)
  
  If LastError_ErrDesc <> "" And blnOnlySetIfBlank = True Then Exit Sub
  LastError_ErrDesc = strErrDesc
  
End Sub

Public Function GetSaveFilter() As String
    GetSaveFilter = "Portable Network Graphics (*.png)|*.png|Joint Photographic Experts Group (*.jpg)|*.jpeg|"
    GetSaveFilter = GetSaveFilter & "Windows or OS/2 Bitmap (*.bmp)|*.bmp|Portable Pixmap (*.ppm)|*.ppm|"
    GetSaveFilter = GetSaveFilter & "Tag Image File Format (*.tiff)|*.tiff|Tagged Image File Format (*.targa)|*.targa"
End Function

Public Function GetSaveFilterFormat(index As Integer) As FREE_IMAGE_FORMAT
    Select Case index
    Case 1
        GetSaveFilterFormat = FIF_PNG
    Case 2
        GetSaveFilterFormat = FIF_JPEG
    Case 3
        GetSaveFilterFormat = FIF_BMP
    Case 4
        GetSaveFilterFormat = FIF_PPM
    Case 5
        GetSaveFilterFormat = FIF_TIFF
    Case 6
        GetSaveFilterFormat = FIF_TARGA
    End Select
End Function

Public Function GetSaveFilterExtension(index As Integer) As String
    Select Case index
    Case 1
        GetSaveFilterExtension = ".png"
    Case 2
        GetSaveFilterExtension = ".jpeg"
    Case 3
        GetSaveFilterExtension = ".bmp"
    Case 4
        GetSaveFilterExtension = ".ppm"
    Case 5
        GetSaveFilterExtension = ".tiff"
    Case 6
        GetSaveFilterExtension = ".targa"
    End Select
End Function

Public Function GetOpenFilter() As String
    GetOpenFilter = "Portable Network Graphics (*.png)|*.png|Joint Photographic Experts Group (*.jpg)|*.jpeg|"
    GetOpenFilter = GetOpenFilter & "Windows or OS/2 Bitmap (*.bmp)|*.bmp|Portable Pixmap (*.ppm)|*.ppm|"
    GetOpenFilter = GetOpenFilter & "Bitmap with run-length encoding (*.rle)|*.rle|Zsoft Paintbrush (*.pcx)|*.pcx|"
    GetOpenFilter = GetOpenFilter & "X11 Bitmap Format (*.xbm)|*.xbm|Kodak PhotoCD (*.pcd)|*.pcd|"
    GetOpenFilter = GetOpenFilter & "Tag Image File Format (*.tiff)|*.tiff|Tagged Image File Format (*.targa)|*.targa|"
    GetOpenFilter = GetOpenFilter & "All files (*.*)|*.*"
End Function

