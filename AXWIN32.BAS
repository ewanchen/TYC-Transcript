Attribute VB_Name = "Module2"
''=============================================================================
''
''  (C) Copyright Axtel Inc. 2000
''
''  Title:      AxWin32.Bas
''
''  Description: This file contains the Basic Language declarations
''               and constants for the useful Windows Api functions.
''
''
''  Version:    1.00
''
''  Date:       24 May 2000
''
''  Author:     Robert E. Payan
''
''-----------------------------------------------------------------------------
''
''  Change log:
''
''     DATE     REV                 DESCRIPTION
''  ----------- ---- ----------------------------------------------------------
''  24 May 2000 1.00 Created from AxTool32.bas.
''                  Robert E. Payan
''=============================================================================
Public Const BI_bitfields = 3&
Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const BITSPIXEL = 12
Public Const DIB_PAL_COLORS = 1
Public Const DIB_PAL_INDICES = 2
Public Const DIB_PAL_LOGINDICES = 4
Public Const DIB_PAL_PHYSINDICES = 2
Public Const DIB_RGB_COLORS = 0
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const INVALID_HANDLE_VALUE = -1
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3

Public Const WHITE_BRUSH = 0
Public Const LTGRAY_BRUSH = 1
Public Const GRAY_BRUSH = 2
Public Const DKGRAY_BRUSH = 3
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
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
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO_EX
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type
Public Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type
Public Type BITMAPCOREHEADER
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type
Public Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As RGBTRIPLE
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharLower1 Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemory01 Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub CopyMemory10 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemory11 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CreateFile1 Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, ByVal Fill As Byte)
Public Declare Sub FillMemory1 Lib "kernel32" Alias "RtlFillMemory" (ByVal dest As Long, ByVal numBytes As Long, ByVal Fill As Byte)
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetBitmapBits1 Lib "gdi32" Alias "GetBitmapBits" (ByVal hBitmap As Long, ByVal dwCount As Long, ByVal lpBits As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetObjectBitmap Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpy1 Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Public Declare Function lstrcpy2 Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Public Declare Function lstrcpyn1 Lib "kernel32" Alias "lstrcpynA" (lpString1 As Any, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function lstrlen1 Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function MessageBoxL Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function MessageBoxS Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function ReadFile1 Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function ReadFile2 Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetDIBitsToDevice_Ex Lib "gdi32" Alias "SetDIBitsToDevice" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByVal Bits As Long, BitsInfo As BITMAPINFO_EX, ByVal wUsage As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function WriteFile1 Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function WriteFile2 Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Sub ZeroMemory1 Lib "kernel32" Alias "RtlZeroMemory" (ByVal dest As Long, ByVal numBytes As Long)
Option Explicit

