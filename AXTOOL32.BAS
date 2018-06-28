Attribute VB_Name = "Module1"
''=============================================================================
''
''  (C) Copyright Axtel Inc. 2004
''
''  Title:      AxTool32.Bas
''
''  Description: This file contains the Basic Language declarations
''               and constants for the Axtel, Inc. Scan Image Tools Library.
''
''
''  Version:    2.42
''
''  Date:       01 July 2004
''
''  Author:     Robert E. Payan
''
''-----------------------------------------------------------------------------
''
''  Change log:
''
''     DATE     REV                 DESCRIPTION
''  ----------- ---- ----------------------------------------------------------
''  01 Jul 2004 2.42 Add CODE_93 write type and AX_BAR_READTYPE_CODE_93 read type.
''                      Robert E. Payan
''  23 Apr 2004 2.41 Add AX_2D_RETURN_PARTIAL flag.
''                      Robert E. Payan
''      10 Mar 2004 2.40 Add AX_DM_FORCE_INVERT flag.
''                      Robert E. Payan
''      25 Feb 2004 2.39 Add AX_BAR_RETURN_FOUR_POINTS and AX_2D_RETURN_FOUR_POINTS flags.
''                      Robert E. Payan
''  18 Feb 2004 2.38 Add AX_BAR_CVT_DISABLE and AX_2D_CVT_DISABLE flags.
''                      Robert E. Payan
''  12 Jan 2004 2.37 Add AX_DM_FASTFIND_DISABLE flag.
''                      Robert E. Payan
''  21 May 2003 2.36 Add AxBarReadFile, AxCodeOneReadFile, AxPdf417ReadFile,
''                      Ax7ReadFile and Ax8ReadFile Api functions.
''                      Robert E. Payan
''  01 May 2003 2.35 Add AxReadGetStartPosition, AxReadGetNextSymbol for all 2D types.
''                      Robert E. Payan
''      02 Apr 2003 2.34 Add AX_QR_M1_DEF for QR Code write types.
''                  Robert E. Payan
''      20 Mar 2003 2.33 Add DIB interface for AxDspk, AxLine, AxSkew and AxStamp.
''                      Add colors for stamp.
''                  Robert E. Payan
''      18 Dec 2002 2.32 Add all Axbmp32, Axjpeg32, and Axtiff32 to the Ax132 API.
''                  Robert E. Payan
''      25 Oct 2002 2.31 Add Ax132, Axbmp32, Axjpeg32, and complete Axtiff32.
''                  Robert E. Payan
''  16 Jul 2002 2.29 Add AX_BAR_DIR_HORIZONTAL, AX_BAR_DIR_VERTICAL,
''                      and AX_BAR_DIR_DIAGONAL.
''                      Enable the use of AX_BAR_DIR_SKEW.
''                      Add AX_2D_DIR_HORIZONTAL and AX_2D_DIR_VERTICAL flags.
''                  Robert E. Payan
''  17 Jan 2002 2.28 Add AX_DM_READ_SMALL.
''                  Robert E. Payan
''  31 Jul 2001 2.27 Add MaxiCode read functions and declarations.
''                  Robert E. Payan
''  03 Jul 2001 2.26 Add UCCEAN_128 and AX_BAR_READTYPE_UCCEAN_128 to
''                      1D barcodes.
''                  Robert E. Payan
''  14 Jun 2001 2.25 Add AX_2D_READ_MODE_BASIC, AX_2D_READ_MODE_EXTENDED.
''                      Add PDF417 read modes and macro modes.
''                  Robert E. Payan
''  27 Feb 2001 2.24 Add AX_BAR_RETURNCHECK.
''                      Robert E. Payan
''  20 Feb 2001 2.23 Add DataMatrix Read and Write functions and declarations.
''                      Robert E. Payan
''  30 Nov 2000 2.22 Add AX_2D_RETURNCORRUPT and AX_BAR_FOUNDCORRUPT for PDF417.
''                  Robert E. Payan
''  28 Aug 2000 2.21 Add dwFlags, CheckBoxRatio, and MinCheckBoxLineLength
''                      members to AXLINE32.
''                  Robert E. Payan
''  21 Jul 2000 2.19 Add Api function for AxBarWriteGetSize.
''                  Robert E. Payan
''  15 Jul 2000 2.18 Add Xmodule to AxBar32. Add AX_BAR_USE_XMOD, AX_BAR_JUSTIFYRIGHT
''                      and AX_BAR_JUSTIFYHCENTER to barcode write flags.
''                  Robert E. Payan
''  31 May 2000 2.20 Add Bar Code Write Types - OCX.
''                  Robert E. Payan
''  27 May 2000 2.19 Change the AxStamp arrays from As String * x to (x) As Byte.
''                  Robert E. Payan
''  25 May 2000 2.18 Change the barcode arrays from As String * 64 to (63) As Byte.
''                  Change AxBarDIB32, AxCodeOneDIB, and AxPdf417DIB to take
''                  Any as the pbmi structure. This enables using the expanded
''                  bitmapinfo structure.
''                  Add AxBarDIB32Ex, AxCodeOneDIBEx, and AXPdf417DIBEx declarations
''                  to take a long for pbmi. This enables the user to use globally
''                  allocated ram for the bitmapinfo structure.
''                  Robert E. Payan
''  24 May 2000 2.17 Remove Windows Api functions and declarations to AxWin32.bas.
''                  Robert E. Payan
''  13 Apr 2000 2.16 Add Api functions to free memory that was allocated by the
''                      Ax5r32 and Ax6r32 DLLs.
''                  Robert E. Payan
''  18 Oct 1999 2.00 Add AX_BAR_USECOLORS, AX_BAR_TRANSPARENT, AX_2D_USECOLORS,
''                                       and AX_2D_TRANSPARENT.
''                                       Add crColorBar and crColorSpace to all barcode structures.
''                                       Add nOutType and nOutGrade to tagAXBARMULTIPLE2D structure.
''                   Robert E. Payan
''  14 Sep 1999 1.27 Add AX_NOTSUPPORTED status.
''                                       Add Password calls for AxBar32, Ax5r32, Ax5w32, Ax6r32, and Ax6w32.
''                   Robert E. Payan
''  23 Jul 1999 1.26 Add DIB AIPs. Add GetSize for CodeOne and PDF417 read.
''                   Robert E. Payan
''  11 Jun 1999 1.25 Add AX_BAR_AVOIDBLOCK flag
''                   Robert E. Payan
''  20 May 1999 1.24 Add cMaxLength to AxBar32 structure
''                   Robert E. Payan
''  08 Apr 1999 1.23 Add AX_TIFF_MODE_COUNTSUBFILES
''                   Robert E. Payan
''  13 Aug 1998 1.22 Add AX_2D_BINARYMODE.
''                   Robert E. Payan
''  03 Feb 1998 1.21 Change Change AxBar32 structure.
''                   Robert E. Payan
''  26 Nov 1997 1.20 Change AxBar2D* to AxCodeOne*.
''                   Robert E. Payan
''  15 Sep 1997 1.10 Change name from AxBar132.dll to Ax5R32.dll.
''                   Change name from AxBar232.dll to Ax5W32.dll.
''                   Robert E. Payan
''  18 Aug 1997 1.00 Created.
''                   Robert E. Payan
''=============================================================================


''
'' Return values
''
Public Const AX_RETURN_NORMAL = 0


''
'' Image structure errors returned by all engines
''
Public Const AX_OPTIONS_POINTER = 10
Public Const AX_OPTIONS_UNITS = 11
Public Const AX_OPTIONS_SIZE = 12
Public Const AX_USAGE_LIMIT = 13
Public Const AX_NOTSUPPORTED = 14

Public Const AX_IMAGE_POINTER = 20
Public Const AX_IMAGE_WIDTH = 21
Public Const AX_IMAGE_HEIGHT = 22
Public Const AX_IMAGE_BUFF_WIDTH = 23
Public Const AX_IMAGE_DATA_POINTER = 24
Public Const AX_IMAGE_BITSPERPIXEL = 25
Public Const AX_IMAGE_PLANES = 26
Public Const AX_IMAGE_DPI = 27
Public Const AX_IMAGE_FILLORDER = 28
Public Const AX_IMAGE_WHITEDOT = 29
Public Const AX_IMAGE_SIZE = 30
Public Const AX_IMAGE_BUFF_HEIGHT = 31
Public Const AX_BITMAP = 32
Public Const AX_BITMAPINFO_POINTER = 33
Public Const AX_FILENAME = 34


''
'' Errors returned by the Barcode engine
''
Public Const AX_MODE = 40
Public Const AX_BAR_MODE = 40
Public Const AX_BAR_TYPE = 41
Public Const AX_BAR_ERRORCHECK = 42
Public Const AX_BAR_TEXTOUT = 43
Public Const AX_BAR_WIDTH = 44
Public Const AX_BAR_HEIGHT = 45
Public Const AX_BAR_TOSMALL = 46
Public Const AX_BAR_STRING = 47
Public Const AX_BAR_NOTFOUND = 48
Public Const AX_BAR_UNITS = 49
Public Const AX_BAR_MULTIPLE = 50
Public Const AX_BAR_PMULTIPLE = 51
Public Const AX_BAR_MULTIPLEMAX = 52
Public Const AX_BAR_GROUP = 53
Public Const AX_BAR_FLAGS = 54
Public Const AX_BAR_FOUNDCORRUPT = 55

''
'' Errors returned by the Despeck/crop engine
''
Public Const AX_DESPECK_MODE = 60
Public Const AX_DSPK_MODE = 61
Public Const AX_DSPK_TYPE = 62
Public Const AX_CROP_MODE = 63
Public Const AX_CROP_UNITS = 64

''
'' Errors returned by the Line Removal engine
''
Public Const AX_LINE_MODE = 66
Public Const AX_LINE_HLENGTH = 67
Public Const AX_LINE_VLENGTH = 68
Public Const AX_LINE_RATIO = 69

''
'' Errors returned by the Skew engine
''
Public Const AX_SKEW_TYPE = 70
Public Const AX_SKEW_ANGLE = 71
Public Const AX_SKEW_MARGIN = 72
Public Const AX_SKEW_SEARCHWIDTH = 73
Public Const AX_SKEW_FINDMODE = 74
Public Const AX_SKEW_FINDUNITS = 75
Public Const AX_SKEW_FINDGRANULARITY = 76
Public Const AX_SKEW_ROTATEMODE = 77
Public Const AX_SKEW_ROATATEUNITS = 78

''
'' Errors returned by the Tiff engine
''
Public Const AX_CLOSEERROR = 80
Public Const AX_CREATEERROR = 81
Public Const AX_OPENERROR = 82
Public Const AX_READERROR = 83
Public Const AX_WRITEERROR = 84
Public Const AX_NOTTIFF = 85
Public Const AX_TIFFEOF = 86
Public Const AX_COMPRESSERROR = 87
Public Const AX_INITERROR = 88
Public Const AX_CONVERTERROR = 89
Public Const AX_STRIPERROR = 90
Public Const AX_TAGERROR = 91
Public Const AX_MEMORYERROR = 92
Public Const AX_UNSUPPORTEDERROR = 93
Public Const AX_MULTIPAGEERROR = 94

Public Const AX_JPEG_NOTJPEG = 150
Public Const AX_JPEG_DZ = 151
Public Const AX_JPEG_QC = 152
Public Const AX_JPEG_NC = 153
Public Const AX_JPEG_DM = 154
Public Const AX_JPEG_DCM = 155
Public Const AX_JPEG_CF = 156
Public Const AX_JPEG_CC = 157
Public Const AX_JPEG_CM = 158
Public Const AX_JPEG_OC = 159
Public Const AX_JPEG_SMOOTH = 160
Public Const AX_JPEG_READMARKERS = 161
Public Const AX_JPEG_READCOMMENT = 162
Public Const AX_JPEG_WRITEMARKERS = 163
Public Const AX_JPEG_WRITECOMMENT = 164



''
'' Processing Modes
''
Public Const AX_MODE_IMAGE = 2

''
'' These macros are used for binary decisions.
''
Public Const AX_NO = 0
Public Const AX_YES = 1

''
'' Bit orders
''
Public Const AX_FIRST_BIT_LEFT = 0
Public Const AX_FIRST_BIT_RIGHT = 1

''
'' White color
''
Public Const AX_WHITE_ZERO = 0
Public Const AX_WHITE_ONE = 1
Public Const AX_WHITE_TAG = &HFF

''
'' Units of measurment in an image
''
Public Const AX_SCANSPIXELS = 0
Public Const AX_INCHESX100 = 1
Public Const AX_MILLIMETERS = 2


''**********************************************************
'' Image Description
''
'' This structure describes the image for all of the engines
''
''**********************************************************
Public Const SizeofAxImage32 = 48&

Type AXIMAGE32STRUCT
    dwType As Long
    dwLineLength As Long
    dwNumLines As Long
    dwBuffWidth As Long
    wPlanes As Integer
    wBitsPerPixel As Integer
    pBits As Long
    dwBuffNumLines As Long
    dwDpiX As Long
    dwDpiY As Long
    dwSize As Long
    hBits As Long
    cFillOrder As Byte
    cWhiteDot As Byte
    Filler As Integer
End Type


''**********************************************************
'' Barcode Section
''
'' This structure describes the barcode parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

''
'' Bar Code Operation Modes
''
Public Const AX_BAR_MODE_READ = 0
Public Const AX_BAR_MODE_WRITE = 1
Public Const AX_BAR_MODE_READMULTIPLE = 2

''
'' Bar Code Write Types - DLL
''
Public Const EAN_13 = 0
Public Const EAN_8 = 1
Public Const UPCA = 2
Public Const UPCE = 3
Public Const CODE_39 = 4
Public Const CODE_128 = 6
Public Const CODE_I25 = 7
Public Const CODA_BAR = 8
Public Const UCCEAN_128 = 9
Public Const CODE_93 = 10
Public Const UNKNOWN = &HFF

''
'' Bar Code Write Types - OCX
''
Public Const EAN_13_OCX = 0
Public Const EAN_8_OCX = 1
Public Const UPCA_OCX = 2
Public Const UPCE_OCX = 3
Public Const CODE_39_OCX = 4
Public Const CODE_128_OCX = 5
Public Const CODE_I25_OCX = 6
Public Const CODA_BAR_OCX = 7

''
'' Bar Code Read Type Masks
''
Public Const AX_BAR_READTYPE_EAN_13 = &H1&
Public Const AX_BAR_READTYPE_EAN_8 = &H2&
Public Const AX_BAR_READTYPE_UPCA = &H4&
Public Const AX_BAR_READTYPE_UPCE = &H8&
Public Const AX_BAR_READTYPE_CODE_39 = &H10&
Public Const AX_BAR_READTYPE_CODE_128 = &H40&
Public Const AX_BAR_READTYPE_CODE_I25 = &H80&
Public Const AX_BAR_READTYPE_CODA_BAR = &H100&
Public Const AX_BAR_READTYPE_UCCEAN_128 = &H200&
Public Const AX_BAR_READTYPE_CODE_93 = &H400&
Public Const AX_BAR_READTYPE_ALL = &HFFFF

''
'' Bar Code Direction
''
Public Const AX_BAR_DIR_LEFT_TO_RIGHT = &H1
Public Const AX_BAR_DIR_RIGHT_TO_LEFT = &H2
Public Const AX_BAR_DIR_TOP_TO_BOTTOM = &H4
Public Const AX_BAR_DIR_BOTTOM_TO_TOP = &H8
Public Const AX_BAR_DIR_SKEW = &H10
Public Const AX_BAR_DIR_HORIZONTAL = &H3
Public Const AX_BAR_DIR_VERTICAL = &HC
Public Const AX_BAR_DIR_DIAGONAL = &H10

''
'' Bar Code Flags
''
Public Const AX_BAR_DESTROYIMAGEOK = &H1&
Public Const AX_BAR_PARTIALSOK = &H2&
Public Const AX_BAR_AVOIDBLOCK = &H4&
Public Const AX_BAR_USE_XMOD = &H8&

Public Const AX_BAR_JUSTIFYRIGHT = &H10&
Public Const AX_BAR_JUSTIFYHCENTER = &H20&

Public Const AX_BAR_USECOLORS = &H200&
Public Const AX_BAR_TRANSPARENT = &H400&
Public Const AX_BAR_RETURNCHECK = &H1000&
Public Const AX_BAR_FLAG_ERRORCHECK = &H2000&
Public Const AX_BAR_CVT_DISABLE = &H20&
Public Const AX_BAR_RETURN_FOUR_POINTS = &H100&

Public Const SizeofAxBarMulti32 = 128&

Type AXBARMULTIPLE32STRUCT
    Status As Long
    OutType As Long
    Duplicate As Long
    OutStartX As Long
    OutStartY As Long
    OutSizeX As Long
    OutSizeY As Long
    Barcode(63) As Byte
    Filler(35) As Byte
End Type

Public Const SizeofAxBar32 = 148&

Type AXBAR32STRUCT
    dwSize As Long
    dwStartX As Long
    dwStartY As Long
    dwSizeX As Long
    dwSizeY As Long
    dwOutStartX As Long
    dwOutStartY As Long
    dwOutSizeX As Long
    dwOutSizeY As Long
    dwReadType As Long
    cDirection As Byte
    cMode As Byte
    cType As Byte
    cUnits As Byte
    cErrorCheck As Byte
    cTextOut As Byte
    cOutType As Byte
    cGranularity As Byte
    cMinLength As Byte
    cMaxLength As Byte
    Filler2 As Integer
    pMultiple As Long
    MultipleMax As Long
    MultipleCount As Long
    WhiteLines As Long
    strBarcode(63) As Byte
    dwFlags As Long
    crColorBar As Long
    crColorSpace As Long
        XModule As Long
End Type

Declare Function AxBar32 Lib "AxBar32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR32STRUCT) As Integer
Declare Function AxBarBitmap32 Lib "AxBar32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR32STRUCT, ByVal hBitmap As Long) As Integer
Declare Function AxBarDIB32 Lib "AxBar32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR32STRUCT, pbmi As Any) As Integer
Declare Function AxBarDIB32Ex Lib "AxBar32" Alias "AxBarDIB32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR32STRUCT, ByVal pbmi As Long) As Integer
Declare Function AxBarInit32 Lib "AxBar32" (pBar As AXBAR32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxBarGetVersion32 Lib "AxBar32" () As Long
Declare Sub AxBarPassword Lib "AxBar32" (pPassword As String)
Declare Function AxBarWriteGetSize Lib "AxBar32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR32STRUCT) As Integer
Declare Function AxBarReadFile Lib "AxBar32" (pBar As AXBAR32STRUCT, pFileName As String, ByVal SubFile As Long) As Integer

''**********************************************************
'' 2Dimensional Barcode Section
''
'' This structure describes the barcode parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

''
'' Code One Write Sub Types
''
Public Const AX_CODE1_DEF = 0
Public Const AX_CODE1_TDEF = 1
Public Const AX_CODE1_SDEF = 2
Public Const AX_CODE1_A = 3
Public Const AX_CODE1_B = 4
Public Const AX_CODE1_C = 5
Public Const AX_CODE1_D = 6
Public Const AX_CODE1_E = 7
Public Const AX_CODE1_F = 8
Public Const AX_CODE1_G = 9
Public Const AX_CODE1_H = 10
Public Const AX_CODE1_T16 = 11
Public Const AX_CODE1_T32 = 12
Public Const AX_CODE1_T48 = 13
Public Const AX_CODE1_S10 = 14
Public Const AX_CODE1_S20 = 15
Public Const AX_CODE1_S30 = 16

''
'' Axtel 2D Bar Code Read Sub Types
''
Public Const AX_CODE1_ATOH = 1&
Public Const AX_CODE1_T = 2&
Public Const AX_CODE1_S = 4&
Public Const AX_CODE1_ALL = &HFFFFFFFF

''
'' Axtel 2D Barcode Read Flags
''
Public Const DESTROYIMAGEOK = 1&
Public Const AX_2D_DESTROYIMAGEOK = 1&
Public Const AX_2D_DIR_LEFTTORIGHT = 2&
Public Const AX_2D_DIR_RIGHTTOLEFT = 4&
Public Const AX_2D_DIR_TOPTOBOTTOM = 8&
Public Const AX_2D_DIR_BOTTOMTOTOP = 16&
Public Const AX_2D_DIR_HORIZONTAL = 6&
Public Const AX_2D_DIR_VERTICAL = 24&
Public Const AX_2D_RETURNCORRUPT = 128&
Public Const AX_2D_CVT_DISABLE = 32&
Public Const AX_2D_RETURN_PARTIAL = &H800&
Public Const AX_2D_RETURN_FOUR_POINTS = &H100&

Public Const AX_2D_READ_MODE_BASIC = 0&
Public Const AX_2D_READ_MODE_EXTENDED = &H1000&

Public Const AX_PDF_READ_MODE_0 = 0&
Public Const AX_PDF_READ_MODE_1 = &H1000&
Public Const AX_PDF_READ_MODE_2 = &H2000&
Public Const AX_PDF_READ_MODE_3_BASIC = &H3000&
Public Const AX_PDF_READ_MODE_3_EXTENDED = &H7000&

Public Const AX_PDF_READ_MACRO_OPTION_0 = &H10000
Public Const AX_PDF_READ_MACRO_OPTION_1 = &H20000
Public Const AX_PDF_READ_MACRO_OPTION_2 = &H40000
Public Const AX_PDF_READ_MACRO_OPTION_3 = &H80000
Public Const AX_PDF_READ_MACRO_OPTION_4 = &H100000
Public Const AX_PDF_READ_MACRO_OPTION_5 = &H200000
Public Const AX_PDF_READ_MACRO_OPTION_6 = &H400000
Public Const AX_PDF_READ_MACRO_OPTION_79AZ = &H800000

''
'' Axtel 2D Barcode Write Flags
''
Public Const AX_2D_TRUNCATE = 1&
Public Const AX_2D_MSGAPPEND = 2&
Public Const AX_2D_INITREADER = 4&
Public Const AX_2D_COLROWASLIMITS = 8&

Public Const AX_2D_JUSTIFYRIGHT = 16&
Public Const AX_2D_JUSTIFYHCENTER = 32&
Public Const AX_2D_JUSTIFYBOTTOM = 64&
Public Const AX_2D_JUSTIFYVCENTER = 128&

Public Const AX_2D_BINARYMODE = 256&

Public Const AX_2D_USECOLORS = &H200&
Public Const AX_2D_TRANSPARENT = &H400&

Public Const SizeofAxBar2DMulti = 44&
Type AXBARMULTIPLE2DSTRUCT
    nOutStatus As Long
    nOutUnits As Long
    nOutType As Byte
    nOutGrade As Byte
    Filler As Integer
    nOutGroup As Long
    nOutSizeofBarcodeData As Long
    nOutStartX As Long
    nOutStartY As Long
    nOutSizeX As Long
    nOutSizeY As Long
    nOutOffsetToNext As Long
    nOutBarcodeData(3) As Byte
End Type

Public Const SizeofAxBar2DRead = 56&
Type AXBAR2DREADSTRUCT
    nSize As Long
    nUnits As Long
    nStartX As Long
    nStartY As Long
    nSizeX As Long
    nSizeY As Long
    dwSearchTypes As Long
    dwFlags As Long
    nMultipleMax As Long
    nOutMultipleCount As Long
    pOutMultiple As Long
    hOutMultiple As Long
    crColorBar As Long
    crColorSpace As Long
End Type

Public Const SizeofAxCodeOneWrite = 64&
Type AXCODEONEWRITESTRUCT
    nSize As Long
    nUnits As Long
    nStartX As Long
    nStartY As Long
    nSizeX As Long
    nSizeY As Long
    dwFlags As Long
    pBuffer As Long
    hBuffer As Long
    nElementX As Long
    nElementY As Long
    nCount As Long
    nType As Long
    nGroup As Long
    crColorBar As Long
    crColorSpace As Long
End Type

Public Const SizeofAxPdf417Write = 68&
Type AXPDF417WRITESTRUCT
    nSize As Long
    nUnits As Long
    nStartX As Long
    nStartY As Long
    nSizeX As Long
    nSizeY As Long
    dwFlags As Long
    pBuffer As Long
    hBuffer As Long
    nBufferCount As Long
    nGroup As Long
    wEccPercent As Integer
    wEccLevel As Integer
    wAspectH As Integer
    wAspectW As Integer
    wModuleAspectRatio As Integer
    wCols As Integer
    wRows As Integer
    XModule As Integer
    crColorBar As Long
    crColorSpace As Long
End Type

Declare Function AxCodeOneRead Lib "Ax5r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT) As Long
Declare Function AxCodeOneReadBitmap Lib "Ax5r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal hBitmap As Long) As Long
Declare Function AxCodeOneReadDIB Lib "Ax5r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, pbmi As Any) As Long
Declare Function AxCodeOneReadDIBEx Lib "Ax5r32" Alias "AxCodeOneReadDIB" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal pbmi As Long) As Long
Declare Function AxCodeOneReadInit Lib "Ax5r32" (pBar As AXBAR2DREADSTRUCT, ByVal dwSize As Long) As Long
Declare Function AxCodeOneReadGetVersion Lib "Ax5r32" () As Long
Declare Sub AxCodeOneReadPassword Lib "Ax5r32" (pPassword As String)
Declare Sub AxCodeOneReadFreeMemory Lib "Ax5r32" (hMem As Long, pMem As Long)
Declare Function AxCodeOneReadGetStartPosition Lib "Ax5r32" (pBar As AXBAR2DREADSTRUCT) As Long
Declare Function AxCodeOneReadGetNextSymbol Lib "Ax5r32" (ByRef pos As Long) As Long
Declare Function AxCodeOneReadFile Lib "Ax5r32" (pBar As AXBAR2DREADSTRUCT, pFileName As String, ByVal SubFile As Long) As Long

Declare Function AxCodeOneWrite Lib "Ax5w32" (pImage As AXIMAGE32STRUCT, pBar As AXCODEONEWRITESTRUCT) As Long
Declare Function AxCodeOneWriteBitmap Lib "Ax5w32" (pImage As AXIMAGE32STRUCT, pBar As AXCODEONEWRITESTRUCT, ByVal hBitmap As Long) As Long
Declare Function AxCodeOneWriteDIB Lib "Ax5w32" (pImage As AXIMAGE32STRUCT, pBar As AXCODEONEWRITESTRUCT, pbmi As Any) As Long
Declare Function AxCodeOneWriteDIBEx Lib "Ax5w32" Alias "AxCodeOneWriteDIB" (pImage As AXIMAGE32STRUCT, pBar As AXCODEONEWRITESTRUCT, ByVal pbmi As Long) As Long
Declare Function AxCodeOneWriteInit Lib "Ax5w32" (pBar As AXCODEONEWRITESTRUCT, ByVal dwSize As Long) As Long
Declare Function AxCodeOneWriteGetVersion Lib "Ax5w32" () As Long
Declare Function AxCodeOneWriteGetSize Lib "Ax5w32" (pImage As AXIMAGE32STRUCT, pBar As AXCODEONEWRITESTRUCT) As Long
Declare Sub AxCodeOneWritePassword Lib "Ax5w32" (pPassword As String)

Declare Function AxPdf417Read Lib "Ax6r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT) As Long
Declare Function AxPdf417ReadBitmap Lib "Ax6r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal hBitmap As Long) As Long
Declare Function AxPdf417ReadDIB Lib "Ax6r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, pbmi As Any) As Long
Declare Function AxPdf417ReadDIBEx Lib "Ax6r32" Alias "AxPdf417ReadDIB" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal pbmi As Long) As Long
Declare Function AxPdf417ReadInit Lib "Ax6r32" (pBar As AXBAR2DREADSTRUCT, ByVal dwSize As Long) As Long
Declare Function AxPdf417ReadGetVersion Lib "Ax6r32" () As Long
Declare Sub AxPdf417ReadPassword Lib "Ax6r32" (pPassword As String)
Declare Sub AxPdf417ReadFreeMemory Lib "Ax6r32" (hMem As Long, pMem As Long)
Declare Function AxPdf417ReadGetStartPosition Lib "Ax6r32" (pBar As AXBAR2DREADSTRUCT) As Long
Declare Function AxPdf417ReadGetNextSymbol Lib "Ax6r32" (ByRef pos As Long) As Long
Declare Function AxPdf417ReadFile Lib "Ax6r32" (pBar As AXBAR2DREADSTRUCT, pFileName As String, ByVal SubFile As Long) As Long

Declare Function AxPdf417Write Lib "Ax6w32" (pImage As AXIMAGE32STRUCT, pBar As AXPDF417WRITESTRUCT) As Long
Declare Function AxPdf417WriteBitmap Lib "Ax6w32" (pImage As AXIMAGE32STRUCT, pBar As AXPDF417WRITESTRUCT, ByVal hBitmap As Long) As Long
Declare Function AxPdf417WriteDIB Lib "Ax6w32" (pImage As AXIMAGE32STRUCT, pBar As AXPDF417WRITESTRUCT, pbmi As Any) As Long
Declare Function AxPdf417WriteDIBEx Lib "Ax6w32" Alias "AxPdf417WriteDIB" (pImage As AXIMAGE32STRUCT, pBar As AXPDF417WRITESTRUCT, ByVal pbmi As Long) As Long
Declare Function AxPdf417WriteInit Lib "Ax6w32" (pBar As AXPDF417WRITESTRUCT, ByVal dwSize As Long) As Long
Declare Function AxPdf417WriteGetVersion Lib "Ax6w32" () As Long
Declare Function AxPdf417WriteGetSize Lib "Ax6w32" (pImage As AXIMAGE32STRUCT, pBar As AXPDF417WRITESTRUCT) As Long
Declare Sub AxPdf417WritePassword Lib "Ax6w32" (pPassword As String)

''
'' **********************************************************
'' DataMatrix Section
''
'' This structure describes the DataMatrix parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
'' **********************************************************
''

''
'' Data Matrix Write Types
''
Public Const AX_DM_DEF = 0
Public Const AX_DM_10x10 = 1
Public Const AX_DM_12x12 = 2
Public Const AX_DM_14x14 = 3
Public Const AX_DM_16x16 = 4
Public Const AX_DM_18x18 = 5
Public Const AX_DM_20x20 = 6
Public Const AX_DM_22x22 = 7
Public Const AX_DM_24x24 = 8
Public Const AX_DM_26x26 = 9
Public Const AX_DM_32x32 = 10
Public Const AX_DM_36x36 = 11
Public Const AX_DM_40x40 = 12
Public Const AX_DM_44x44 = 13
Public Const AX_DM_48x48 = 14
Public Const AX_DM_52x52 = 15
Public Const AX_DM_64x64 = 16
Public Const AX_DM_72x72 = 17
Public Const AX_DM_80x80 = 18
Public Const AX_DM_88x88 = 19
Public Const AX_DM_96x96 = 20
Public Const AX_DM_104x104 = 21
Public Const AX_DM_120x120 = 22
Public Const AX_DM_132x132 = 23
Public Const AX_DM_144x144 = 24
Public Const AX_DM_8x18 = 25
Public Const AX_DM_8x32 = 26
Public Const AX_DM_12x26 = 27
Public Const AX_DM_12x36 = 28
Public Const AX_DM_16x36 = 29
Public Const AX_DM_16x48 = 30

Public Const AX_DM_WRITE_RECTANGLE = &H800

''
'' Data Matrix Read Types.
''
Public Const AX_DM_READ_SQUARE = &H10
Public Const AX_DM_READ_RECTANGLE = &H20
Public Const AX_DM_READ_SMALL = &H40

''
'' Data Matrix Read Flags
''
Public Const AX_DM_FASTFIND_DISABLE = &H2000
Public Const AX_DM_FORCE_INVERT = &H400


Type AX7WRITESTRUCT
    nSize As Long
    nUnits As Long
    nStartX As Long
    nStartY As Long
    nSizeX As Long
    nSizeY As Long
    dwFlags As Long
    pBuffer As Long
    hBuffer As Long
    nBufferCount As Long
    GroupNumber As Byte
    GroupTotal As Byte
    FileIDLo As Byte
    FileIDHi As Byte
        Type As Integer
    XModule As Integer
    crColorBar As Long
    crColorSpace As Long
End Type

Public Const SizeofAx7Write = 56&

''
'' **********************************************************
'' Function Declarations
'' **********************************************************
''
Declare Function Ax7Write Lib "Ax7w32" (pImage As AXIMAGE32STRUCT, pBar As AX7WRITESTRUCT) As Long
Declare Function Ax7WriteBitmap Lib "Ax7w32" (pImage As AXIMAGE32STRUCT, pBar As AX7WRITESTRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax7WriteDIB Lib "Ax7w32" (pImage As AXIMAGE32STRUCT, pBar As AX7WRITESTRUCT, pbmi As Any) As Long
Declare Function Ax7WriteDIBEx Lib "Ax7w32" Alias "Ax7WriteDIB" (pImage As AXIMAGE32STRUCT, pBar As AX7WRITESTRUCT, ByVal pbmi As Long) As Long
Declare Function Ax7WriteInit Lib "Ax7w32" (pBar As AX7WRITESTRUCT, ByVal dwSize As Long) As Long
Declare Function Ax7WriteGetVersion Lib "Ax7w32" () As Long
Declare Function Ax7WriteGetSize Lib "Ax7w32" (pImage As AXIMAGE32STRUCT, pBar As AX7WRITESTRUCT) As Long
Declare Sub Ax7WritePassword Lib "Ax7w32" (pPassword As String)

Declare Function Ax7Read Lib "Ax7r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT) As Long
Declare Function Ax7ReadBitmap Lib "Ax7r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax7ReadDIB Lib "Ax7r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, pbmi As Any) As Long
Declare Function Ax7ReadDIBEx Lib "Ax7r32" Alias "Ax7ReadDIB" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal pbmi As Long) As Long
Declare Function Ax7ReadInit Lib "Ax7r32" (pBar As AXBAR2DREADSTRUCT, ByVal dwSize As Long) As Long
Declare Function Ax7ReadGetVersion Lib "Ax7r32" () As Long
Declare Sub Ax7ReadPassword Lib "Ax7r32" (pPassword As String)
Declare Sub Ax7ReadFreeMemory Lib "Ax7r32" (hMem As Long, pMem As Long)
Declare Function Ax7ReadGetStartPosition Lib "Ax7r32" (pBar As AXBAR2DREADSTRUCT) As Long
Declare Function Ax7ReadGetNextSymbol Lib "Ax7r32" (ByRef pos As Long) As Long
Declare Function Ax7ReadFile Lib "Ax7r32" (pBar As AXBAR2DREADSTRUCT, pFileName As String, ByVal SubFile As Long) As Long

'' **********************************************************
'' QR Code Section
''
'' This structure describes the QR Code parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
'' **********************************************************

''
'' QR Code Write Types
''
Public Const AX_QR_DEF = 0
Public Const AX_QR_M2_1 = 1
Public Const AX_QR_M2_2 = 2
Public Const AX_QR_M2_3 = 3
Public Const AX_QR_M2_4 = 4
Public Const AX_QR_M2_5 = 5
Public Const AX_QR_M2_6 = 6
Public Const AX_QR_M2_7 = 7
Public Const AX_QR_M2_8 = 8
Public Const AX_QR_M2_9 = 9
Public Const AX_QR_M2_10 = 10
Public Const AX_QR_M2_11 = 11
Public Const AX_QR_M2_12 = 12
Public Const AX_QR_M2_13 = 13
Public Const AX_QR_M2_14 = 14
Public Const AX_QR_M2_15 = 15
Public Const AX_QR_M2_16 = 16
Public Const AX_QR_M2_17 = 17
Public Const AX_QR_M2_18 = 18
Public Const AX_QR_M2_19 = 19
Public Const AX_QR_M2_20 = 20
Public Const AX_QR_M2_21 = 21
Public Const AX_QR_M2_22 = 22
Public Const AX_QR_M2_23 = 23
Public Const AX_QR_M2_24 = 24
Public Const AX_QR_M2_25 = 25
Public Const AX_QR_M2_26 = 26
Public Const AX_QR_M2_27 = 27
Public Const AX_QR_M2_28 = 28
Public Const AX_QR_M2_29 = 29
Public Const AX_QR_M2_30 = 30
Public Const AX_QR_M2_31 = 31
Public Const AX_QR_M2_32 = 32
Public Const AX_QR_M2_33 = 33
Public Const AX_QR_M2_34 = 34
Public Const AX_QR_M2_35 = 35
Public Const AX_QR_M2_36 = 36
Public Const AX_QR_M2_37 = 37
Public Const AX_QR_M2_38 = 38
Public Const AX_QR_M2_39 = 39
Public Const AX_QR_M2_40 = 40
Public Const AX_QR_M1_1 = 41
Public Const AX_QR_M1_2 = 42
Public Const AX_QR_M1_3 = 43
Public Const AX_QR_M1_4 = 44
Public Const AX_QR_M1_5 = 45
Public Const AX_QR_M1_6 = 46
Public Const AX_QR_M1_7 = 47
Public Const AX_QR_M1_8 = 48
Public Const AX_QR_M1_9 = 49
Public Const AX_QR_M1_10 = 50
Public Const AX_QR_M1_11 = 51
Public Const AX_QR_M1_12 = 52
Public Const AX_QR_M1_13 = 53
Public Const AX_QR_M1_14 = 54
Public Const AX_QR_M1_DEF = 55

Public Const AX_QR_ECC_L = 0
Public Const AX_QR_ECC_M = 1
Public Const AX_QR_ECC_Q = 2
Public Const AX_QR_ECC_H = 3

Type AX8WRITESTRUCT
    nSize As Long
    nUnits As Long
    nStartX As Long
    nStartY As Long
    nSizeX As Long
    nSizeY As Long
    dwFlags As Long
    pBuffer As Long
    hBuffer As Long
    nBufferCount As Long
    GroupNumber As Long
    GroupTotal As Long
        Type As Long
        EccLevel As Long
    XModule As Long
    crColorBar As Long
    crColorSpace As Long
End Type

Public Const SizeofAx8Write = 68&


'' **********************************************************
'' Function Declarations
'' **********************************************************
Declare Function Ax8Write Lib "Ax8w32" (pImage As AXIMAGE32STRUCT, pBar As AX8WRITESTRUCT) As Long
Declare Function Ax8WriteBitmap Lib "Ax8w32" (pImage As AXIMAGE32STRUCT, pBar As AX8WRITESTRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax8WriteDIB Lib "Ax8w32" (pImage As AXIMAGE32STRUCT, pBar As AX8WRITESTRUCT, pbmi As Any) As Long
Declare Function Ax8WriteDIBEx Lib "Ax8w32" Alias "Ax8WriteDIB" (pImage As AXIMAGE32STRUCT, pBar As AX8WRITESTRUCT, ByVal pbmi As Long) As Long
Declare Function Ax8WriteInit Lib "Ax8w32" (pBar As AX8WRITESTRUCT, ByVal dwSize As Long) As Long
Declare Function Ax8WriteGetVersion Lib "Ax8w32" () As Long
Declare Function Ax8WriteGetSize Lib "Ax8w32" (pImage As AXIMAGE32STRUCT, pBar As AX8WRITESTRUCT) As Long
Declare Sub Ax8WritePassword Lib "Ax8w32" (pPassword As String)

Declare Function Ax8Read Lib "Ax8r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT) As Long
Declare Function Ax8ReadBitmap Lib "Ax8r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax8ReadDIB Lib "Ax8r32" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, pbmi As Any) As Long
Declare Function Ax8ReadDIBEx Lib "Ax8r32" Alias "Ax8ReadDIB" (pImage As AXIMAGE32STRUCT, pBar As AXBAR2DREADSTRUCT, ByVal pbmi As Long) As Long
Declare Function Ax8ReadInit Lib "Ax8r32" (pBar As AXBAR2DREADSTRUCT, ByVal dwSize As Long) As Long
Declare Function Ax8ReadGetVersion Lib "Ax8r32" () As Long
Declare Sub Ax8ReadPassword Lib "Ax8r32" (pPassword As String)
Declare Sub Ax8ReadFreeMemory Lib "Ax8r32" (hMem As Long, pMem As Long)
Declare Function Ax8ReadGetStartPosition Lib "Ax8r32" (pBar As AXBAR2DREADSTRUCT) As Long
Declare Function Ax8ReadGetNextSymbol Lib "Ax8r32" (ByRef pos As Long) As Long
Declare Function Ax8ReadFile Lib "Ax8r32" (pBar As AXBAR2DREADSTRUCT, pFileName As String, ByVal SubFile As Long) As Long

''**********************************************************
'' Despeck/Crop Section
''
'' This structure describes the despeck and crop parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

''
'' Despeck Types
''
Public Const AX_DSPK_5x5_DESPECK = 0
Public Const AX_DSPK_5x5_REPAIR = 1
Public Const AX_DSPK_3x3_DESPECK = 2
Public Const AX_DSPK_3x3_REPAIR = 3
Public Const AX_DSPK_5x5_REPAIR_HQ = 4
Public Const AX_DSPK_3x3_REPAIR_HQ = 5


Public Const SizeofAxDspk32 = 28&

Type AXDSPK32STRUCT
    dwSize As Long
    dwCropTopMargin As Long
    dwCropLeftMargin As Long
    dwCropRightMargin As Long
    dwCropBottomMargin As Long
    cMode As Byte
    cDspkMode As Byte
    cDspkType As Byte
    cCropMode As Byte
    cCropUnits As Byte
    Filler1 As Byte
    Filler2 As Integer
End Type

Declare Function AxDspk32 Lib "AxDspk32" (pImage As AXIMAGE32STRUCT, pDspk As AXDSPK32STRUCT) As Integer
Declare Function AxDspkBitmap32 Lib "AxDspk32" (pImage As AXIMAGE32STRUCT, pDspk As AXDSPK32STRUCT, ByVal hBitmap As Long) As Integer
Declare Function AxDspkDIB32 Lib "AxDspk32" (pImage As AXIMAGE32STRUCT, pDspk As AXDSPK32STRUCT, pbmi As Any) As Long
Declare Function AxDspkDIB32Ex Lib "AxDspk32" Alias "AxDspkDIB32" (pImage As AXIMAGE32STRUCT, pDspk As AXDSPK32STRUCT, ByVal pbmi As Long) As Long
Declare Function AxDspkInit32 Lib "AxDspk32" (pDspk As AXDSPK32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxDspkGetVersion32 Lib "AxDspk32" () As Long

''**********************************************************
'' Line Removal Section
''
'' This structure describes the line removal parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

''
'' Line Removal Modes
''
Public Const AX_LINE_MODE_BOTH = 0
Public Const AX_LINE_MODE_HORIZ = 1
Public Const AX_LINE_MODE_VERTI = 2

''
'' Flag Bits
''
Public Const AX_LINE_CHECKBOX = 1
Public Const AX_LINE_REMOVAL = 2

Public Const SizeofAxLine32 = 32&

Type AXLINE32STRUCT
    dwSize As Long
    dwMinHLength As Long
    dwMinVLength As Long
    cMode As Byte
    cConnect As Byte
    Filler As Integer
    fRatio As Single
        CheckBoxRatio As Single
        dwFlags As Long
        MinCheckBoxLineLength As Long
End Type

Declare Function AxLine32 Lib "AxLine32" (pImage As AXIMAGE32STRUCT, pLine As AXLINE32STRUCT) As Integer
Declare Function AxLineBitmap32 Lib "AxLine32" (pImage As AXIMAGE32STRUCT, pLine As AXLINE32STRUCT, ByVal hBitmap As Long) As Integer
Declare Function AxLineDIB32 Lib "AxLine32" (pImage As AXIMAGE32STRUCT, pLine As AXLINE32STRUCT, pbmi As Any) As Long
Declare Function AxLineDIB32Ex Lib "AxLine32" Alias "AxLineDIB32" (pImage As AXIMAGE32STRUCT, pLine As AXLINE32STRUCT, ByVal pbmi As Long) As Long
Declare Function AxLineInit32 Lib "AxLine32" (pLine As AXLINE32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxLineGetVersion32 Lib "AxLine32" () As Long

''**********************************************************
'' Skew Section
''
'' This structure describes the skew parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

Public Const AX_SKEW_FIND_ANGLE_AND_ROTATE = 0
Public Const AX_SKEW_FIND_ANGLE = 1
Public Const AX_SKEW_ROTATE = 2
Public Const AX_SKEW_FIND_ANGLE_AND_ROTATE1 = 3
Public Const AX_SKEW_REGISTER = 4
Public Const AX_SKEW_FINDANGLE_REGISTER = 5
Public Const AX_SKEW_REGISTER_ROTATE = 6
Public Const AX_SKEW_FINDANGLE_REGISTER_ROTATE = 7

Public Const AX_SKEW_TYPE_FINDANGLE = &H1
Public Const AX_SKEW_TYPE_ROTATE = &H2
Public Const AX_SKEW_TYPE_REGISTER = &H4


Public Const AX_SKEW_ROTATE_BOTH = 0
Public Const AX_SKEW_ROTATE_HORIZ = 1
Public Const AX_SKEW_ROTATE_VERTI = 2

Public Const AX_SKEW_FIND_SPECIAL = 0
Public Const AX_SKEW_FIND_LEFT = 1

Public Const SizeofAxSkew32 = 36&

Type AXSKEW32STRUCT
    dwSize As Long
    dwAngleX As Long
    dwAngleY As Long
    dwTopX As Long
    dwTopY As Long
    dwFindMargin As Long
    dwFindSearchWidth As Long

    cSkewType As Byte
    cFindMode As Byte
    cFindUnits As Byte
    cFindGranularity As Byte
    cRotateMode As Byte
    cRotateUnits As Byte
    Filler As Integer
End Type

Declare Function AxSkew32 Lib "AxSkew32" (pImage As AXIMAGE32STRUCT, pSkew As AXSKEW32STRUCT) As Integer
Declare Function AxSkewBitmap32 Lib "AxSkew32" (pImage As AXIMAGE32STRUCT, pSkew As AXSKEW32STRUCT, ByVal hBitmap As Long) As Integer
Declare Function AxSkewDIB32 Lib "AxSkew32" (pImage As AXIMAGE32STRUCT, pSkew As AXSKEW32STRUCT, pbmi As Any) As Long
Declare Function AxSkewDIB32Ex Lib "AxSkew32" Alias "AxSkewDIB32" (pImage As AXIMAGE32STRUCT, pSkew As AXSKEW32STRUCT, ByVal pbmi As Long) As Long
Declare Function AxSkewInit32 Lib "AxSkew32" (pSkew As AXSKEW32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxSkewGetVersion32 Lib "AxSkew32" () As Long


''**********************************************************
'' Stamp Section
''
'' This structure describes the stamp parameters.
'' It must be used in conjunction with the image structure
'' described above.
''
''**********************************************************

Public Const AX_STAMP_MODEWRITE = 1&
Public Const AX_STAMP_MODEFINDSIZE = 2&
Public Const AX_STAMP_MODEFINDPOS = 4&
Public Const AX_STAMP_MODEUSECOLORS = 8&

Public Const AX_STAMP_TRANSPARENT = &H1
Public Const AX_STAMP_BOLD = &H2
Public Const AX_STAMP_UNDERLINE = &H4
Public Const AX_STAMP_ITALIC = &H8
Public Const AX_STAMP_STRIKEOUT = &H10

Public Const AX_STAMP_FORMAT0 = 0
Public Const AX_STAMP_FORMAT1 = 1
Public Const AX_STAMP_FORMAT2 = 2

Public Const LF_FACESIZE = 32
Public Const SizeofAxStmp32 = 164&

Type AXSTMP32STRUCT
    Size As Long
    ModeFlags As Long
    StartX As Long
    StartY As Long
    SizeX As Long
    SizeY As Long
    SearchStartX As Long
    SearchStartY As Long
    SearchSizeX As Long
    SearchSizeY As Long
    Units As Long
    FormatIndex As Long
    FontFlags As Long
    FontPointSize As Long
    Number As Long
    sName(63) As Byte
    sFontName(LF_FACESIZE - 1) As Byte
        crColorBar As Long
        crColorSpace As Long
End Type

Declare Function AxStamp32 Lib "AxStmp32" (pImage As AXIMAGE32STRUCT, pStmp As AXSTMP32STRUCT) As Integer
Declare Function AxStampBitmap32 Lib "AxStmp32" (pImage As AXIMAGE32STRUCT, pStmp As AXSTMP32STRUCT, ByVal hBitmap As Long) As Integer
Declare Function AxStampDIB32 Lib "AxStmp32" (pImage As AXIMAGE32STRUCT, pStmp As AXSTMP32STRUCT, pbmi As Any) As Long
Declare Function AxStampDIB32Ex Lib "AxStmp32" Alias "AxStampDIB32" (pImage As AXIMAGE32STRUCT, pStmp As AXSTMP32STRUCT, ByVal pbmi As Long) As Long
Declare Function AxStampInit32 Lib "AxStmp32" (pStmp As AXSTMP32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxStampGetVersion32 Lib "AxStmp32" () As Long

''**********************************************************
'' BMP File Read/Write Section
''**********************************************************

Public Const AX_BMP_MODE_WRITE = 0
Public Const AX_BMP_MODE_READ = 1
Public Const AX_BMP_MODE_READ_HEADER = 2

Type AXBMP32STRUCT
    pFileName As Long
    Mode As Long
    dwFlags As Long
    dwSize As Long
End Type


Declare Function AxBmpGetVersion Lib "Ax1bmp32" () As Long

Declare Function AxBmpReadInit Lib "Ax1bmp32" (pBmp As AXBMP32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxBmpRead Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT) As Long
Declare Function AxBmpReadBitmap Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByRef hBitmap As Long) As Long
Declare Function AxBmpReadDIB Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, pbmi As Any) As Long
Declare Function AxBmpReadDIBEx Lib "Ax1bmp32" Alias "AxBmpReadDIB" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal pbmi As Long) As Long
Declare Sub AxBmpReadFreeMemory Lib "Ax1bmp32" (hMem As Long, pMem As Long)

Declare Function AxBmpWriteInit Lib "Ax1bmp32" (pBmp As AXBMP32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxBmpWrite Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT) As Long
Declare Function AxBmpWriteBitmap Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal hBitmap As Long) As Long
Declare Function AxBmpWriteDIB Lib "Ax1bmp32" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, pbmi As Any) As Long
Declare Function AxBmpWriteDIBEx Lib "Ax1bmp32" Alias "AxBmpWriteDIB" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal pbmi As Long) As Long

''**********************************************************
'' Jpeg File Read/Write Section
''**********************************************************

Public Const AX_JPEG_MODE_WRITE = 0
Public Const AX_JPEG_MODE_READ = 1
Public Const AX_JPEG_MODE_READ_HEADER = 2

Public Const AX_JPEG_1TO1 = 0
Public Const AX_JPEG_1TO2 = 1
Public Const AX_JPEG_1TO4 = 2
Public Const AX_JPEG_1TO8 = 3

Public Const AX_JPEG_COLOR_TRUE = 0
Public Const AX_JPEG_COLOR_MAPPED = 1
Public Const AX_JPEG_COLOR_GRAYSCALE = 2

Public Const AX_JPEG_DITHER_NONE = 0
Public Const AX_JPEG_DITHER_ORDERED = 1
Public Const AX_JPEG_DITHER_FS = 2

Public Const AX_JPEG_DECOMP_SLOWINT = 0
Public Const AX_JPEG_DECOMP_FASTINT = 1
Public Const AX_JPEG_DECOMP_FLOAT = 2

Type AXJPEG32STRUCT
    pFileName As Long
    Mode As Long
    dwFlags As Long
    dwSize As Long

        ReadOutputColorSpace As Long
        ReadDecompressZoom As Long
        ReadQuantizeColors As Long
        ReadNumberOfColors As Long
        ReadDitherMethod As Long
        ReadDecompressMethod As Long
        ReadSmoothEnable As Long

        WriteCompressFact As Long
        WriteCompressColors As Long
        WriteCompressMethod As Long
        WriteCompressType As Long
        WriteSmoothFactor As Long
        WriteOptimize As Long
        WriteInColorSpace As Long

        pMarker_App0 As Long
        pMarker_App1 As Long
        pMarker_App2 As Long
        pMarker_App3 As Long
        pMarker_App4 As Long
        pMarker_App5 As Long
        pMarker_App6 As Long
        pMarker_App7 As Long
        pMarker_App8 As Long
        pMarker_App9 As Long
        pMarker_App10 As Long
        pMarker_App11 As Long
        pMarker_App12 As Long
        pMarker_App13 As Long
        pMarker_App14 As Long
        pComment As Long

        ReadMarker_App0Length As Integer
        ReadMarker_App1Length As Integer
        ReadMarker_App2Length As Integer
        ReadMarker_App3Length As Integer
        ReadMarker_App4Length As Integer
        ReadMarker_App5Length As Integer
        ReadMarker_App6Length As Integer
        ReadMarker_App7Length As Integer
        ReadMarker_App8Length As Integer
        ReadMarker_App9Length As Integer
        ReadMarker_App10Length As Integer
        ReadMarker_App11Length As Integer
        ReadMarker_App12Length As Integer
        ReadMarker_App13Length As Integer
        ReadMarker_App14Length As Integer
        ReadCommentLength As Integer

End Type



Declare Function AxJpegGetVersion Lib "Ax1jpg32" () As Long

Declare Function AxJpegReadInit Lib "Ax1jpg32" (pJpeg As AXJPEG32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxJpegRead Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT) As Long
Declare Function AxJpegReadBitmap Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByRef hBitmap As Long) As Long
Declare Function AxJpegReadDIB Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, pbmi As Any) As Long
Declare Function AxJpegReadDIBEx Lib "Ax1jpg32" Alias "AxJpegReadDIB" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal pbmi As Long) As Long
Declare Sub AxJpegReadFreeMemory Lib "Ax1jpg32" (hMem As Long, pMem As Long)

Declare Function AxJpegWriteInit Lib "Ax1jpg32" (pJpeg As AXJPEG32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxJpegWrite Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT) As Long
Declare Function AxJpegWriteBitmap Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal hBitmap As Long) As Long
Declare Function AxJpegWriteDIB Lib "Ax1jpg32" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, pbmi As Any) As Long
Declare Function AxJpegWriteDIBEx Lib "Ax1jpg32" Alias "AxJpegWriteDIB" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal pbmi As Long) As Long



''**********************************************************
'' Tiff File Read/Write Section
''**********************************************************

''
'' These macros are used for the output compression types.
''
Public Const AX_OUT_IMAGE = 0
Public Const AX_OUT_G4 = 1
Public Const AX_OUT_G32D = 3
Public Const AX_OUT_G31D = 7
Public Const AX_OUT_TIFG3 = 8
Public Const AX_OUT_PAKBITS = 9

''
'' These macros are used for eol_type.
''
Public Const AX_EOL_BIT = 0
Public Const AX_EOL_BYTE = 1

''
'' Scaling types
''
Public Const AX_SCALE_NONE = 0
Public Const AX_SCALE_YES = 1

''
'' Image Contrast
''
Public Const AX_CONTRAST_DEFAULT = 0
Public Const AX_CONTRAST_1250 = 1
Public Const AX_CONTRAST_2500 = 2
Public Const AX_CONTRAST_3750 = 3
Public Const AX_CONTRAST_5000 = 4
Public Const AX_CONTRAST_6250 = 5
Public Const AX_CONTRAST_7500 = 6
Public Const AX_CONTRAST_8750 = 7

''
'' Image Orientaion
''
Public Const AX_TIFF_ORIENT_DEFAULT = 0
Public Const AX_TIFF_ORIENT_1 = 1
Public Const AX_TIFF_ORIENT_2 = 2
Public Const AX_TIFF_ORIENT_3 = 3
Public Const AX_TIFF_ORIENT_4 = 4
Public Const AX_TIFF_ORIENT_5 = 5
Public Const AX_TIFF_ORIENT_6 = 6
Public Const AX_TIFF_ORIENT_7 = 7
Public Const AX_TIFF_ORIENT_8 = 8

''
'' Tiff NewSubfileType bits
''
Public Const AX_TIFF_NS_RESOLUTION = &H1&
Public Const AX_TIFF_NS_SINGLEPAGE = &H2&
Public Const AX_TIFF_NS_TRANSMASK = &H4&

Public Const AX_TIFF_MODE_WRITE = 0
Public Const AX_TIFF_MODE_READ = 1
Public Const AX_TIFF_MODE_READ_HEADER = 2
Public Const AX_TIFF_MODE_COUNTSUBFILES = 3

Public Const SIATEL_DISABLE = 0
Public Const SIATEL_WHITE_ONE = 1
Public Const SIATEL_WHITE_ZERO = 2

Public Const SizeofAxTiff32 = 112&

Type AXTIFF32STRUCT
    pFileName As Long
    pDocumentName As Long
    pImageDescription As Long
    pMake As Long
    pModel As Long
    pPageName As Long
    pArtist As Long
    pCopyright As Long

    ReadDocumentNameLength As Long
    ReadImageDescriptionLength As Long
    ReadMakeLength As Long
    ReadModelLength As Long
    ReadPageNameLength As Long
    ReadArtistLength As Long
    ReadCopyrightLength As Long

    ReadNewSubfileType As Long
    ReadSubfile As Integer
    ReadPageNumber As Integer
    ReadPageTotal As Integer
    ReadHorizScale As Integer
    ReadVertiScale As Integer
    ReadScale As Byte
    ReadContrast As Byte

    WriteNewSubfileType As Long
    WriteMultiTiffEnable As Integer
    WritePageNumber As Integer
    WritePageTotal As Integer
    WriteCompType As Byte
    WriteEolType As Byte
    WriteG3KFactor As Long
    WriteRowsPerStrip As Long

    dwSize As Long
    PhotometricControl As Long
    Mode As Byte
    WriteFillOrder As Byte

    ReadTargetOrientation As Integer
    ReadOrientation As Integer
    WriteOrientation As Integer

End Type

Declare Function AxTiffGetVersion32 Lib "Ax1tif32" () As Long

Declare Function AxTiffReadInit Lib "Ax1tif32" (pTiff As AXTIFF32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxTiffRead Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT) As Long
Declare Function AxTiffReadBitmap Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByRef hBitmap As Long) As Long
Declare Function AxTiffReadDIB Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, pbmi As Any) As Long
Declare Function AxTiffReadDIBEx Lib "Ax1tif32" Alias "AxTiffReadDIB" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal pbmi As Long) As Long
Declare Sub AxTiffReadFreeMemory Lib "Ax1tif32" (hMem As Long, pMem As Long)

Declare Function AxTiffWriteInit Lib "Ax1tif32" (pTiff As AXTIFF32STRUCT, ByVal dwSize As Long) As Long
Declare Function AxTiffWrite Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT) As Long
Declare Function AxTiffWriteBitmap Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal hBitmap As Long) As Long
Declare Function AxTiffWriteDIB Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, pbmi As Any) As Long
Declare Function AxTiffWriteDIBEx Lib "Ax1tif32" Alias "AxTiffWriteDIB" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal pbmi As Long) As Long

Declare Function AxTiff32 Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT) As Integer
Declare Function AxTiffBitmap32 Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByRef hBitmap As Long) As Integer
Declare Function AxTiffInit32 Lib "Ax1tif32" (pTiff As AXTIFF32STRUCT, ByVal dwSize As Long) As Integer
Declare Function AxTiffInitImage32 Lib "Ax1tif32" (pImage As AXIMAGE32STRUCT, dwSize As Long) As Integer

''**********************************************************
'' Generic File Read/Write Section
''**********************************************************

Public Const AX_MODE_WRITE = 0
Public Const AX_MODE_READ = 1
Public Const AX_MODE_READ_HEADER = 2
Public Const AX_MODE_READ_COUNTSUBFILES = 3

Public Const AX_1_MODE_WRITE = 0
Public Const AX_1_MODE_READ = 1
Public Const AX_1_MODE_READ_HEADER = 2
Public Const AX_1_MODE_READ_COUNTSUBFILES = 3

Type AX132STRUCT
    pFileName As Long
    Mode As Long
    dwFlags As Long
    dwSize As Long
    ReadNewSubfileType As Long
    ReadSubfile As Long
    ReadPageNumber As Long
    ReadPageTotal As Long
End Type


Declare Function Ax1GetVersion Lib "Ax132" () As Long

Declare Function Ax1ReadInit Lib "Ax132" (pAx1 As AX132STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1Read Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT) As Long
Declare Function Ax1ReadBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, ByRef hBitmap As Long) As Long
Declare Function Ax1ReadDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, pbmi As Any) As Long
Declare Function Ax1ReadDIBEx Lib "Ax132" Alias "Ax1ReadDIB" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, ByVal pbmi As Long) As Long
Declare Sub Ax1ReadFreeMemory Lib "Ax132" (hMem As Long, pMem As Long)

Declare Function Ax1WriteInit Lib "Ax132" (pAx1 As AX132STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1Write Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT) As Long
Declare Function Ax1WriteBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax1WriteDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, pbmi As Any) As Long
Declare Function Ax1WriteDIBEx Lib "Ax132" Alias "Ax1WriteDIB" (pImage As AXIMAGE32STRUCT, pAx1 As AX132STRUCT, ByVal pbmi As Long) As Long


Declare Function Ax1BmpGetVersion Lib "Ax132" () As Long

Declare Function Ax1BmpReadInit Lib "Ax132" (pBmp As AXBMP32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1BmpRead Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT) As Long
Declare Function Ax1BmpReadBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByRef hBitmap As Long) As Long
Declare Function Ax1BmpReadDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, pbmi As Any) As Long
Declare Function Ax1BmpReadDIBEx Lib "Ax132" Alias "Ax1BmpReadDIB" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal pbmi As Long) As Long
Declare Sub Ax1BmpReadFreeMemory Lib "Ax132" (hMem As Long, pMem As Long)

Declare Function Ax1BmpWriteInit Lib "Ax132" (pBmp As AXBMP32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1BmpWrite Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT) As Long
Declare Function Ax1BmpWriteBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax1BmpWriteDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, pbmi As Any) As Long
Declare Function Ax1BmpWriteDIBEx Lib "Ax132" Alias "Ax1BmpWriteDIB" (pImage As AXIMAGE32STRUCT, pBmp As AXBMP32STRUCT, ByVal pbmi As Long) As Long

Declare Function Ax1JpegGetVersion Lib "Ax132" () As Long

Declare Function Ax1JpegReadInit Lib "Ax132" (pJpeg As AXJPEG32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1JpegRead Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT) As Long
Declare Function Ax1JpegReadBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByRef hBitmap As Long) As Long
Declare Function Ax1JpegReadDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, pbmi As Any) As Long
Declare Function Ax1JpegReadDIBEx Lib "Ax132" Alias "Ax1JpegReadDIB" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal pbmi As Long) As Long
Declare Sub Ax1JpegReadFreeMemory Lib "Ax132" (hMem As Long, pMem As Long)

Declare Function Ax1JpegWriteInit Lib "Ax132" (pJpeg As AXJPEG32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1JpegWrite Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT) As Long
Declare Function Ax1JpegWriteBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax1JpegWriteDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, pbmi As Any) As Long
Declare Function Ax1JpegWriteDIBEx Lib "Ax132" Alias "Ax1JpegWriteDIB" (pImage As AXIMAGE32STRUCT, pJpeg As AXJPEG32STRUCT, ByVal pbmi As Long) As Long

Declare Function Ax1TiffGetVersion Lib "Ax132" () As Long

Declare Function Ax1TiffReadInit Lib "Ax132" (pTiff As AXTIFF32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1TiffRead Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT) As Long
Declare Function Ax1TiffReadBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByRef hBitmap As Long) As Long
Declare Function Ax1TiffReadDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, pbmi As Any) As Long
Declare Function Ax1TiffReadDIBEx Lib "Ax132" Alias "Ax1TiffReadDIB" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal pbmi As Long) As Long
Declare Sub Ax1TiffReadFreeMemory Lib "Ax132" (hMem As Long, pMem As Long)

Declare Function Ax1TiffWriteInit Lib "Ax132" (pTiff As AXTIFF32STRUCT, ByVal dwSize As Long) As Long
Declare Function Ax1TiffWrite Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT) As Long
Declare Function Ax1TiffWriteBitmap Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal hBitmap As Long) As Long
Declare Function Ax1TiffWriteDIB Lib "Ax132" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, pbmi As Any) As Long
Declare Function Ax1TiffWriteDIBEx Lib "Ax132" Alias "Ax1TiffWriteDIB" (pImage As AXIMAGE32STRUCT, pTiff As AXTIFF32STRUCT, ByVal pbmi As Long) As Long


Option Explicit
