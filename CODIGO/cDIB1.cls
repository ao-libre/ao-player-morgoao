VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'*  VB file:   cDIB.cls... by Ray Mercer
'*  created:   12/1999 by Ray Mercer
'*  uploaded:  2/2000
'*  modified:  2/25/2000 by Ray Mercer
'*             Patrick Pasteels pointed out a bug in my code
'*             -fixed: ReDim m_memBitmapInfo(0 To 39) now correctly equals 40 bytes
'*
'*
'*  Copyright (C) 1999 - 2000 Ray Mercer.  All rights reserved.
'*  Latest version can be downloaded from http://www.shrinkwrapvb.com
'****************************************************************
Option Explicit

Private Const BMP_MAGIC_COOKIE As Integer = 19778 'this is equivalent to ascii string "BM"
'//BITMAP DEFINES (from mmsystem.h)
Private Type BITMAPFILEHEADER '14 bytes
        bfType As Integer '"magic cookie" - must be "BM"
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
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

Private Type RGBQUAD
    red As Byte
    green As Byte
    blue As Byte
    Reserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0

Private Const BI_RGB  As Long = 0&

Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal HDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long 'handle
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long 'Pointer to mem
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long 'BOOL
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef Src As Any, ByVal dwLen As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal HDC As Long) As Long

Private Const HEAP_ZERO_MEMORY As Long = &H8

Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER

' Handle to the current DIBSection:
Private m_hDIb As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

Private wDC As Long
Private cDC As Long
Private cBM As Long


Public Function CreateFromFile(ByVal FileName As String) As Boolean
    Dim hFile As Long
        
    If Not ExistFile(FileName) Then
        MsgBox "File does not exist:" & vbCrLf & FileName, vbCritical, App.Title
        Exit Function
    End If
        
    hFile = FreeFile()
    
    '<====ERROR TRAP ON
    On Error Resume Next
    Open FileName For Binary Access Read As #hFile
    If Err Then
        If Err.number = 70 Then
            MsgBox "File is locked - cannot access:" & vbCrLf & FileName, vbCritical, App.Title
        Else
            MsgBox Err.Description, vbInformation, App.Title
        End If
        Exit Function 'assume file was not opened
    End If
    On Error GoTo 0
    '====>ERROR TRAP OFF
    
    'OK, file is opened - now for the real algorithm...
    Get #hFile, , m_bfh 'get the BITMAPFILEHEADER this identifies the bitmap

    If m_bfh.bfType <> BMP_MAGIC_COOKIE Then 'this is not a BMP file
        MsgBox "File is not a supported bitmap format:" & vbCrLf & FileName, vbInformation, App.Title
        Close #hFile
        Exit Function
    Else
        'now get the info header
        Get #hFile, Len(m_bfh) + 1, m_bih 'start at the 15th byte
        
        'now get the bitmap bits
        ReDim m_memBits(0 To m_bih.biSizeImage - 1)
        Get #hFile, m_bfh.bfOffBits + 1, m_memBits
        
        'and BitmapInfo variable-length UDT
        ReDim m_memBitmapInfo(0 To m_bfh.bfOffBits - 14) 'don't need first 14 bytes (fileinfo)
        Get #hFile, Len(m_bfh) + 1, m_memBitmapInfo
        
        Close #hFile   'Close file
    End If
    
    CreateFromFile = True 'indicate success
End Function

Private Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (Err = 0)
End Function

Public Property Get BitCount() As Long
    BitCount = m_bih.biBitCount
End Property

Public Property Get Height() As Long
    Height = m_bih.biHeight
End Property

Public Property Get Width() As Long
    Width = m_bih.biWidth
End Property

Public Property Get Compression() As Long
    Compression = m_bih.biCompression
End Property

Public Property Get SizeInfoHeader() As Long
    SizeInfoHeader = m_bih.biSize
End Property

Public Property Get SizeImage() As Long
    SizeImage = m_bih.biSizeImage
End Property

Public Property Get Planes() As Long
    Planes = m_bih.biPlanes
End Property

Public Property Get ClrImportant() As Long
    ClrImportant = m_bih.biClrImportant
End Property

Public Property Get ClrUsed() As Long
    ClrUsed = m_bih.biClrUsed
End Property

Public Property Get XPPM() As Long
    XPPM = m_bih.biXPelsPerMeter
End Property

Public Property Get YPPM() As Long
    YPPM = m_bih.biYPelsPerMeter
End Property

Public Property Get FileType() As Long
    FileType = m_bfh.bfType
End Property

Public Property Get SizeFileHeader() As Long
    SizeFileHeader = m_bfh.bfSize
End Property

Public Property Get BitOffset() As Long
    BitOffset = m_bfh.bfOffBits
End Property

Public Property Get PointerToBits() As Long
    PointerToBits = VarPtr(m_memBits(0))
'    Debug.Print Hex(PointerToBits)
End Property

Public Property Get PointerToBitmapInfo() As Long
    PointerToBitmapInfo = VarPtr(m_memBitmapInfo(0))
End Property

Public Property Get SizeBitmapInfo() As Long
    SizeBitmapInfo = UBound(m_memBitmapInfo()) + 1
End Property

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0
        End If
    End If
End Function

Public Sub LoadPictureBlt( _
                ByVal lhDC As Long, _
                Optional ByVal lSrcLeft As Long = 0, _
                Optional ByVal lSrcTop As Long = 0, _
                Optional ByVal lSrcWidth As Long = -1, _
                Optional ByVal lSrcHeight As Long = -1, _
                Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
                
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub
Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteDC m_hDC
    End If
    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property
Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Function CleanUp()
Dim res As Long
Dim tmpBIH As BITMAPINFOHEADER
Dim tmpBFH As BITMAPFILEHEADER
    
    If wDC Then
        If ReleaseDC(frmMain.hWnd, wDC) = 0 Then
            Debug.Print "No se puede releasear el wDC"
        End If
    End If
    If cBM Then
        If DeleteObject(cBM) = 0 Then
            Debug.Print "Error al deletear cBM"
        End If
    End If
    If cDC Then
        If DeleteDC(cDC) = 0 Then
            Debug.Print "Error al deletear cDC"
        End If
    End If
    If m_hDC Then
        If DeleteDC(m_hDC) = 0 Then
            Debug.Print "Error al deletear el m_hDC"
        End If
    End If
    If m_hBmpOld Then
        If DeleteObject(m_hBmpOld) = 0 Then
            Debug.Print "Error al deletear m_hBmpOld"
        End If
    End If
End Function

'Budi
Public Function createMagically(ByVal Width As Long, ByVal Height As Long) As Boolean
Dim ret As Long
Dim tBMP As BITMAP

Call CleanUp

wDC = GetDC(frmMain.hWnd)
If wDC = 0 Then createMagically = False: Exit Function

m_hDC = CreateCompatibleDC(wDC)
If m_hDC = 0 Then createMagically = False: Exit Function

cBM = CreateCompatibleBitmap(wDC, Width, Height)
If cBM = 0 Then createMagically = False: Exit Function

m_hBmpOld = SelectObject(m_hDC, cBM)
If m_hBmpOld = 0 Then createMagically = False: Exit Function

Call BitBlt(m_hDC, 0, 0, Width, Height, wDC, 0, 0, &HCC0020)
Call GetObjectAPI(cBM, Len(tBMP), tBMP)

m_bfh.bfOffBits = 54 'same for all 24-bit bitmaps
m_bfh.bfReserved1 = 0 'reserved, set to 0
m_bfh.bfReserved2 = 0 'again, reserved, set to 0
m_bfh.bfSize = Len(m_bfh) + Len(m_bih) + (3 * tBMP.bmHeight * tBMP.bmWidth)
m_bfh.bfType = BMP_MAGIC_COOKIE 'BMP type
    
m_bih.biBitCount = 24 ' 24bpp
m_bih.biClrImportant = 0 ' number of important colout table entries
m_bih.biClrUsed = 0 ' number of colour table entries that are used
m_bih.biCompression = 0 ' no compression
m_bih.biHeight = tBMP.bmHeight ' the height of the bitmap
m_bih.biPlanes = tBMP.bmPlanes ' the number of planes in the bitmap
m_bih.biSize = Len(m_bih) ' the size of this header
m_bih.biSizeImage = 3 * (tBMP.bmHeight * tBMP.bmWidth) ' the size of the image data
m_bih.biWidth = tBMP.bmWidth ' the width of the bitmap
m_bih.biXPelsPerMeter = 0 ' pixels per meter along the x axis, not required
m_bih.biYPelsPerMeter = 0 ' pixels per meter along the y axis, not required

Call SelectObject(m_hDC, m_hBmpOld)

ReDim m_memBits(3 * (tBMP.bmHeight * tBMP.bmWidth) - 1) As Byte
m_tBI.bmiHeader = m_bih
ret = GetDIBits(m_hDC, cBM, 0, m_bih.biHeight, m_memBits(0), m_tBI, DIB_RGB_COLORS)

If ret = 0 Then
    Call CleanUp
    createMagically = False
Else
    ReDim m_memBitmapInfo(Len(m_tBI.bmiHeader)) As Byte
    Call CopyMemory(m_memBitmapInfo(0), m_tBI, Len(m_tBI.bmiHeader))
    createMagically = True
End If

End Function
