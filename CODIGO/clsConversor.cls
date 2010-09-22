VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConversor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long 'returns fourcc
Private Declare Function VideoForWindowsVersion Lib "msvfw32.dll" () As Long
Private Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
Private Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pFile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
Private Declare Function AVIFileCreateStream Lib "avifil32.dll" Alias "AVIFileCreateStreamA" (ByVal pFile As Long, ByRef ppavi As Long, ByRef psi As AVI_STREAM_INFO) As Long
Private Declare Function AVISaveOptions Lib "avifil32.dll" (ByVal hWnd As Long, ByVal uiFlags As Long, ByVal nStreams As Long, ByRef ppavi As Long, ByRef ppOptions As Long) As Long
Private Declare Function AVISave Lib "avifil32.dll" Alias "AVISaveVA" (ByVal szFile As String, ByVal pclsidHandler As Long, ByVal lpfnCallback As Long, ByVal nStreams As Long, ByRef ppaviStream As Long, ByRef ppCompOptions As Long) As Long
Private Declare Function AVISaveOptionsFree Lib "avifil32.dll" (ByVal nStreams As Long, ByRef ppOptions As Long) As Long
Private Declare Function AVIMakeCompressedStream Lib "avifil32.dll" (ByRef ppsCompressed As Long, ByVal psSource As Long, ByRef lpOptions As AVI_COMPRESS_OPTIONS, ByVal pclsidHandler As Long) As Long
Private Declare Function AVIStreamSetFormat Lib "avifil32.dll" (ByVal pavi As Long, ByVal lPos As Long, ByRef lpFormat As Any, ByVal cbFormat As Long) As Long
Private Declare Function AVIStreamWrite Lib "avifil32.dll" (ByVal pavi As Long, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long, ByVal dwFlags As Long, ByRef plSampWritten As Long, ByRef plBytesWritten As Long) As Long
Private Declare Function AVIStreamReadFormat Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lPos As Long, ByVal lpFormatBuf As Long, ByRef sizeBuf As Long) As Long
Private Declare Function AVIStreamRead Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long, ByRef pBytesWritten As Long, ByRef pSamplesWritten As Long) As Long
Private Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef bih As Any) As Long
Private Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, ByVal lPos As Long) As Long
Private Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long
Private Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pFile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Private Declare Function AVIMakeFileFromStreams Lib "avifil32.dll" (ByRef ppfile As Long, ByVal nStreams As Long, ByVal pAVIStreamArray As Long) As Long
Private Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
Private Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Private Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
Private Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Private Declare Function AVIStreamClose Lib "avifil32.dll" Alias "AVIStreamRelease" (ByVal pavi As Long) As Long 'ULONG
Private Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pFile As Long) As Long
Private Declare Function AVIFileClose Lib "avifil32.dll" (ByVal pFile As Long) As Long

Private Declare Function AVIMakeStreamFromClipboard Lib "avifil32.dll" (ByVal cfFormat As Long, ByVal hprivate As Long, ByRef ppstream As Long) As Long
Private Declare Function AVIPutFileOnClipboard Lib "avifil32.dll" (ByVal pAVIFile As Long) As Long
Private Declare Function AVIGetFromClipboard Lib "avifil32.dll" (ByRef ppAVIFile As Long) As Long
Private Declare Function AVIClearClipboard Lib "avifil32.dll" () As Long

Private Declare Sub AVIFileInit Lib "avifil32.dll" ()

Private Declare Sub AVIFileExit Lib "avifil32.dll" ()

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function privateAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long


Private Declare Function privateFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Type waveFormat
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
End Type

Private Type PCMWAVEFORMAT
    wf As waveFormat
    wBitsPerSample As Integer
End Type

Private Type AVI_RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type BITMAPFILEHEADER '14 bytes
  bfType As Integer
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

Private Type BITMAPINFOHEADER_MJPEG '68 bytes
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
   biExtDataOffset As Long
   JPEGSize As Long
   JPEGProcess As Long
   JPEGColorSpaceID As Long
   JPEGBitsPerSample As Long
   JPEGHSubSampling As Long
   JPEGVSubSampling As Long
End Type
        
Private Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount  As Long
    dwFormatChangeCount As Long
    szName As String * 64
End Type

Private Type AVI_FILE_INFO
    dwMaxBytesPerSecond As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

Private Type AVI_COMPRESS_OPTIONS
    fccType As Long
    fccHandler As Long
    dwKeyFrameEvery As Long
    dwQuality As Long
    dwBytesPerSecond As Long
    dwFlags As Long
    lpFormat As Long
    cbFormat As Long
    lpParms As Long
    cbParms As Long
    dwInterleaveEvery As Long
End Type

Private Const AVIERR_OK As Long = 0&
Private Const OF_READWRITE As Long = &H2
Private Const AVIIF_KEYFRAME As Long = &H10
Private Const streamtypeVIDEO As Long = 1935960438
Private Const streamtypeAUDIO As Long = 1935963489
Private Const streamtypeMIDI As Long = 1935960429
Private Const streamtypeTEXT As Long = 1937012852
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
'To use this function, need 2 files
'existing avi video file(without sound)
'existing wav audio file

Private Const SEVERITY_ERROR    As Long = &H80000000
Private Const FACILITY_ITF      As Long = &H40000
Private Const AVIERR_BASE       As Long = &H4000

Private Const AVIERR_BADFLAGS    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 105) '-2147205015
Private Const AVIERR_BADPARAM    As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 106) '-2147205014
Private Const AVIERR_BADSIZE     As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 107) '-2147205013
Private Const AVIERR_USERABORT   As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 198) '-2147204922

Private Const AVIFILEINFO_HASINDEX         As Long = &H10
Private Const AVIFILEINFO_MUSTUSEINDEX     As Long = &H20
Private Const AVIFILEINFO_ISINTERLEAVED    As Long = &H100
Private Const AVIFILEINFO_WASCAPTUREFILE   As Long = &H10000
Private Const AVIFILEINFO_COPYRIGHTED      As Long = &H20000

Private Const AVIFILECAPS_CANREAD          As Long = &H1
Private Const AVIFILECAPS_CANWRITE         As Long = &H2
Private Const AVIFILECAPS_ALLKEYFRAMES     As Long = &H10
Private Const AVIFILECAPS_NOCOMPRESSION    As Long = &H20

Private Const AVICOMPRESSF_INTERLEAVE     As Long = &H1           '// interleave
Private Const AVICOMPRESSF_DATARATE       As Long = &H2           '// use a data rate
Private Const AVICOMPRESSF_KEYFRAMES      As Long = &H4           '// use keyframes
Private Const AVICOMPRESSF_VALID          As Long = &H8           '// has valid data?

Private Const OF_READ  As Long = &H0
Private Const OF_WRITE As Long = &H1
Private Const OF_SHARE_DENY_WRITE As Long = &H20
Private Const OF_CREATE As Long = &H1000

Private Const DIB_RGB_COLORS  As Long = 0    '/* color table in RGBs */
Private Const DIB_PAL_COLORS  As Long = 1    '/* color table in palette indices */

Private Const BI_RGB          As Long = 0
Private Const BI_RLE8         As Long = 1
Private Const BI_RLE4         As Long = 2
Private Const BI_BITFIELDS    As Long = 3

Private Const AVIGETFRAMEF_BESTDISPLAYFMT  As Long = 1

Private Const ICMF_CHOOSE_KEYFRAME           As Long = &H1     '// show KeyFrame Every box
Private Const ICMF_CHOOSE_DATARATE           As Long = &H2     '// show DataRate box
Private Const ICMF_CHOOSE_PREVIEW            As Long = &H4     '// allow expanded preview dialog
Private Const ICMF_CHOOSE_ALLCOMPRESSORS     As Long = &H8     '// don't only show those that
                                                               '// can handle the input format
'bltbit constant
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Declare Function SetRect Lib "user32.dll" (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef Src As Any, ByVal dwLen As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Const HEAP_ZERO_MEMORY As Long = &H8

Private outputVideoFileName  As String
Private ps As Long 'ptr PAVISTREAM
Private psCompressed As Long 'ptr PAVISTREAM

Public Converting As Boolean
Public Count As Long
Public pFile As Long
'Public ps As Long

Private Type tDimension
    nWidth As Long
    nHeight As Long
End Type
Private tScreen As tDimension
Private initialBMP As cDib
Private mBMP As cDib
Private strhdr As AVI_STREAM_INFO
Private opts As AVI_COMPRESS_OPTIONS
Private pOpts As Long, res As Long, hdcc As Long

Private currentFrame As Long
Private lastFrame As Long

Public Sub initConversor(ByVal videoFileName As String)
Dim H As Long, W As Long
    Call AVIFileInit

'    Open the file for writing
    res = AVIFileOpen(pFile, videoFileName, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set mBMP = New cDib
    
    'FRMMAIN DE MIERDA.
    frmMain.ScaleMode = vbPixels
    H = frmMain.ScaleX(frmMain.Height, vbTwips, vbPixels)
    W = frmMain.ScaleX(frmMain.Width, vbTwips, vbPixels)
    
    tScreen.nHeight = H
    tScreen.nWidth = W

    If Not mBMP.createMagically(tScreen.nWidth, tScreen.nHeight) Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'    Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)               '// stream type video
        .fccHandler = 0&                                        '// default AVI handler
        .dwScale = 1
        .dwRate = 15                                            '// fps
        .dwSuggestedBufferSize = mBMP.SizeImage                 '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, mBMP.Width, mBMP.Height)   '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pFile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    'or ICMF_CHOOSE_PREVIEW or ICMF_CHOOSE_ALLCOMPRESSORS
    res = AVISaveOptions(frmConnect.hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
      Call AVISaveOptionsFree(1, pOpts)
      GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
'    With BI
'        .biBitCount = mBMP.BitCount
'        .biClrImportant = mBMP.ClrImportant
'        .biClrUsed = mBMP.ClrUsed
'        .biCompression = mBMP.Compression
'        .biHeight = mBMP.Height
'        .biWidth = mBMP.Width
'        .biPlanes = mBMP.Planes
'        .biSize = mBMP.SizeInfoHeader
'        .biSizeImage = mBMP.SizeImage
'        .biXPelsPerMeter = mBMP.XPPM
'        .biYPelsPerMeter = mBMP.YPPM
'    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal mBMP.PointerToBitmapInfo, mBMP.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    Converting = True

Exit Sub
error:
'   Now close the file
    Set mBMP = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pFile <> 0) Then Call AVIFileRelease(pFile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
      MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
End Sub

Public Sub addFrame()
    Dim ret As Long, out1 As Long, out2 As Long
    
    ret = mBMP.createMagically(tScreen.nWidth, tScreen.nHeight)
    
    If ret Then
        ret = AVIStreamWrite(psCompressed, lastFrame, 1, mBMP.PointerToBits, mBMP.SizeImage, AVIIF_KEYFRAME, out1, out2)
        Debug.Print "Frame " & lastFrame & "-> bytes: " & out2 & " - " & "samples: " & out1
    Else
        Debug.Print "Error creando frame " & lastFrame & "-> bytes: " & out2 & " - " & "samples: " & out1
    End If

    If ret <> AVIERR_OK Then MsgBox "err"

    lastFrame = lastFrame + 1
End Sub

Public Sub closeFile()
If (ps <> 0) Then Call AVIStreamRelease(ps)
If (psCompressed <> 0) Then Call AVIStreamRelease(psCompressed)
If (pFile <> 0) Then Call AVIFileRelease(pFile)
Call AVIFileExit
End Sub

Public Property Get getFrame() As Long
getFrame = currentFrame
End Property

Public Sub incrementFrame()
currentFrame = currentFrame + 1
End Sub

Public Sub resetFrame()
currentFrame = 0
End Sub

Public Sub openAVI()
Dim res As Long
End Sub

Public Sub resetFlags()
Dim res As Long

Call AVIStreamRelease(psCompressed)

res = AVIStreamSetFormat(psCompressed, lastFrame + 1, ByVal initialBMP.PointerToBitmapInfo, initialBMP.SizeBitmapInfo)

If (res <> AVIERR_OK) Then
    MsgBox "Error al setear el formato"
    Call closeFile
End If

Call openAVI
End Sub

'Private Sub AddAudioStream(ByVal AVIFilePath As String, ByVal WAVFilePath As String) 'AddAudioStream(Destination AVI FilePath, Source WAV FilePath)
'    On Error GoTo ErrHandler
'    Dim StreamInfo As AVI_STREAM_INFO
'    Dim StreamFormat As PCMWAVEFORMAT
'    Dim StreamLength As Long
'    Dim WaveData As Long
'    Dim AVIFile As Long
'    Dim AudioFile As Long
'    Dim AudioStream As Long
'    Dim AVIStream As Long
'    Call AVIFileInit 'Initialize the AVI library.
'
'
'    If AVIFileOpen(AVIFile, AVIFilePath, OF_READWRITE, ByVal 0&) = AVIERR_OK Then 'open the avi file
'
'
'        If AVIFileOpen(AudioFile, WAVFilePath, OF_READWRITE, ByVal 0&) = AVIERR_OK Then 'open the wave file
'
'            If AVIFileGetStream(AudioFile, AudioStream, streamtypeAUDIO, 0) <> AVIERR_OK Then GoTo ErrHandler  'get the audio stream
'
'            If AVIStreamInfo(AudioStream, StreamInfo, Len(StreamInfo)) <> AVIERR_OK Then GoTo ErrHandler 'read the stream's header information
'
'            AVIStreamReadFormat AudioStream, 0, StreamFormat, Len(StreamFormat) 'read the stream's format data
'
'            StreamLength = AVIStreamLength(AudioStream) * StreamInfo.dwSampleSize 'get the length of the stream
'
'            WaveData = privateAlloc(GPTR, StreamLength) 'get pointer To the wave data
'
'            AVIStreamRead AudioStream, 0, StreamLength, WaveData, StreamLength, 0, 0 'read audio data from the stream
'
'            AVIFileCreateStream AVIFile, AVIStream, StreamInfo 'create new stream
'
'            AVIStreamSetFormat AVIStream, 0, StreamFormat, Len(StreamFormat) 'set the format of new stream
'
'            AVIStreamWrite AVIStream, 0, StreamLength, WaveData, StreamLength, AVIIF_KEYFRAME, 0, 0   'copy the raw wave data To new stream
'
'            privateFree WaveData 'release the wave data pointer
'
'            AVIStreamRelease (AudioStream)
'
'            AVIStreamRelease (AVIStream)
'
'            AVIFileRelease AudioFile
'
'            AVIFileRelease AVIFile
'
'            Call AVIFileExit 'close the AVI library
'        End If
'    End If
'ErrHandler:
'
'End Sub
