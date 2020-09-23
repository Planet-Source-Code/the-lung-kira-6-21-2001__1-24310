Attribute VB_Name = "wingdi"
Option Explicit


Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long


Public Const DRIVERVERSION = 0
Public Const TECHNOLOGY = 2
Public Const HORZSIZE = 4
Public Const VERTSIZE = 6
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const BITSPIXEL = 12
Public Const PLANES = 14
Public Const NUMBRUSHES = 16
Public Const NUMPENS = 18
Public Const NUMMARKERS = 20
Public Const NUMFONTS = 22
Public Const NUMCOLORS = 24
Public Const PDEVICESIZE = 26
Public Const CURVECAPS = 28
Public Const LINECAPS = 30
Public Const POLYGONALCAPS = 32
Public Const TEXTCAPS = 34
Public Const CLIPCAPS = 36
Public Const RASTERCAPS = 38
Public Const ASPECTX = 40
Public Const ASPECTY = 42
Public Const ASPECTXY = 44
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const SIZEPALETTE = 104
Public Const NUMRESERVED = 106
Public Const COLORRES = 108
Public Const PHYSICALWIDTH = 110
Public Const PHYSICALHEIGHT = 111
Public Const PHYSICALOFFSETX = 112
Public Const PHYSICALOFFSETY = 113
Public Const SCALINGFACTORX = 114
Public Const SCALINGFACTORY = 115
Public Const VREFRESH = 116
Public Const DESKTOPVERTRES = 117
Public Const DESKTOPHORZRES = 118
Public Const BLTALIGNMENT = 119
Public Const SHADEBLENDCAPS = 120
Public Const COLORMGMTCAPS = 121

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = &H1
Public Const DISPLAY_DEVICE_MULTI_DRIVER = &H2
Public Const DISPLAY_DEVICE_PRIMARY_DEVICE = &H4
Public Const DISPLAY_DEVICE_MIRRORING_DRIVER = &H8
Public Const DISPLAY_DEVICE_VGA_COMPATIBLE = &H10
Public Const DISPLAY_DEVICE_REMOVABLE = &H20
Public Const DISPLAY_DEVICE_MODESPRUNED = &H8000000
Public Const DISPLAY_DEVICE_REMOTE = &H4000000
Public Const DISPLAY_DEVICE_DISCONNECT = &H2000000

Public Const DM_ORIENTATION = &H1
Public Const DM_PAPERSIZE = &H2
Public Const DM_PAPERLENGTH = &H4
Public Const DM_PAPERWIDTH = &H8
Public Const DM_SCALE = &H10
Public Const DM_POSITION = &H20
Public Const DM_NUP = &H40
Public Const DM_COPIES = &H100
Public Const DM_DEFAULTSOURCE = &H200
Public Const DM_PRINTQUALITY = &H400
Public Const DM_COLOR = &H800
Public Const DM_DUPLEX = &H1000
Public Const DM_YRESOLUTION = &H2000
Public Const DM_TTOPTION = &H4000
Public Const DM_COLLATE = &H8000
Public Const DM_FORMNAME = &H10000
Public Const DM_LOGPIXELS = &H20000
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_ICMMETHOD = &H800000
Public Const DM_ICMINTENT = &H1000000
Public Const DM_MEDIATYPE = &H2000000
Public Const DM_DITHERTYPE = &H4000000
Public Const DM_PANNINGWIDTH = &H8000000
Public Const DM_PANNINGHEIGHT = &H10000000

Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64


Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long

    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Public Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type
