VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrtDevMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const DM_PAPERSIZE = &H2
Private Const DM_PAPERLENGTH = &H4
Private Const DM_PAPERWIDTH = &H8

Private Type RawDataType
    RGB As String * 47
End Type

Private Type DevModeType
    strDeviceName As String * 16
    
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    
    intOrientation As Integer
    intPaperSize_mm As Integer
    intPaperLength_mm As Integer
    intPaperWidth_mm As Integer
    intScale As Integer
    
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    
    intYResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName As String * 16
    lngPad As Long
    
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type

Dim tDevMode As DevModeType

Public Sub LoadData(ByVal vNewValue As Variant)
    Dim rawdata As RawDataType
    rawdata.RGB = vNewValue
    LSet tDevMode = rawdata
End Sub

Public Function ToString()
    Dim rdt As RawDataType
    
    tDevMode.lngFields = tDevMode.lngFields Or DM_PAPERSIZE Or _
                DM_PAPERLENGTH Or DM_PAPERWIDTH
    tDevMode.intPaperSize_mm = 256  'ユーザ定義型を指定する
   
    LSet rdt = tDevMode
    
    ToString = rdt.RGB

End Function

Public Property Get DeviceName() As Variant
    Dim idx As Integer
    Dim v As Variant
    idx = InStrB(1, tDevMode.strDeviceName, StrConv(Chr(0), vbFromUnicode))
    v = MidB(tDevMode.strDeviceName, 1, (idx - 1))
    DeviceName = StrConv(v, vbUnicode)
End Property
Public Property Let DeviceName(ByVal vNewValue As Variant)
    Dim v As String
    v = StrConv(vNewValue, vbFromUnicode)
    tDevMode.strDeviceName = v & StrConv(Chr(0), vbFromUnicode)
End Property

Public Property Get FormName() As Variant
    Dim idx As Integer
    Dim v As String
    
    v = StrConv(MidB(tDevMode.strFormName, 1), vbUnicode)
    idx = InStr(1, v, Chr(0))
    FormName = Left(v, idx - 1)
End Property
Public Property Let FromName(ByVal vNewValue As String)
    Dim v As String
    v = StrConv(vNewValue, vbFromUnicode)
    tDevMode.strFormName = v & StrConv(Chr(0), vbFromUnicode)
End Property

Public Property Get SpecVersion() As Variant
    SpecVersion = tDevMode.intSpecVersion
End Property
Public Property Let SpecVersion(ByVal vNewValue As Variant)
    tDevMode.intSpecVersion = vNewValue
End Property

Public Property Get DriverVersion() As Variant
    DriverVersion = tDevMode.intDriverVersion
End Property
Public Property Let DriverVersion(ByVal vNewValue As Variant)
    tDevMode.intDriverVersion = vNewValue
End Property

Public Property Get Size() As Variant
    Size = tDevMode.intSize
End Property
Public Property Let Size(ByVal vNewValue As Variant)
    tDevMode.intSize = vNewValue
End Property

Public Property Get DriverExtra() As Variant
    DriverExtra = tDevMode.intDriverExtra
End Property
Public Property Let DriverExtra(ByVal vNewValue As Variant)
    tDevMode.intDriverExtra = vNewValue
End Property

Public Property Get Fields() As Variant
    Fields = tDevMode.lngFields
End Property
Public Property Let Fields(ByVal vNewValue As Variant)
    tDevMode.lngFields = vNewValue
End Property

Public Property Get Orientation() As Variant
    Orientation = tDevMode.intOrientation
End Property
Public Property Let Orientation(ByVal vNewValue As Variant)
    tDevMode.intOrientation = vNewValue
End Property

Public Property Get PaperSize_mm() As Variant
    PaperSize_mm = tDevMode.intPaperSize_mm
End Property
Public Property Let PaperSize_mm(ByVal vNewValue As Variant)
    tDevMode.intPaperSize_mm = vNewValue
End Property

Public Property Get PaperLength_mm() As Variant
    PaperLength_mm = tDevMode.intPaperLength_mm
End Property
Public Property Let PaperLength_mm(ByVal vNewValue As Variant)
    tDevMode.intPaperLength_mm = vNewValue
End Property

Public Property Get PaperWidth_mm() As Variant
    PaperWidth_mm = tDevMode.intPaperWidth_mm
End Property
Public Property Let PaperWidth_mm(ByVal vNewValue As Variant)
    tDevMode.intPaperWidth_mm = vNewValue
End Property

Public Property Get Scales() As Variant
    Scales = tDevMode.intScale
End Property
Public Property Let Scales(ByVal vNewValue As Variant)
    tDevMode.intScale = vNewValue
End Property

Public Property Get Copies() As Variant
    Copies = tDevMode.intCopies
End Property
Public Property Let Copies(ByVal vNewValue As Variant)
    tDevMode.intCopies = vNewValue
End Property

Public Property Get DefaultSource() As Variant
    DefaultSource = tDevMode.intDefaultSource
End Property
Public Property Let DefaultSource(ByVal vNewValue As Variant)
    tDevMode.intDefaultSource = vNewValue
End Property

Public Property Get PrintQuality() As Variant
    PrintQuality = tDevMode.intPrintQuality
End Property
Public Property Let PrintQuality(ByVal vNewValue As Variant)
    tDevMode.intPrintQuality = vNewValue
End Property

Public Property Get Color() As Variant
    Color = tDevMode.intColor
End Property
Public Property Let Color(ByVal vNewValue As Variant)
    tDevMode.intColor = vNewValue
End Property

Public Property Get Duplex() As Variant
    Duplex = tDevMode.intDuplex
End Property
Public Property Let Duplex(ByVal vNewValue As Variant)
    tDevMode.intDuplex = vNewValue
End Property

Public Property Get YResolution() As Variant
    YResolution = tDevMode.intYResolution
End Property
Public Property Let YResolution(ByVal vNewValue As Variant)
    tDevMode.intYResolution = vNewValue
End Property

Public Property Get TTOption() As Variant
    TTOption = tDevMode.intTTOption
End Property
Public Property Let TTOption(ByVal vNewValue As Variant)
    tDevMode.intTTOption = vNewValue
End Property

Public Property Get Collate() As Variant
    Collate = tDevMode.intCollate
End Property
Public Property Let Collate(ByVal vNewValue As Variant)
    tDevMode.intCollate = vNewValue
End Property

Public Property Get Pad() As Variant
    Pad = tDevMode.lngPad
End Property
Public Property Let Pad(ByVal vNewValue As Variant)
    tDevMode.lngPad = vNewValue
End Property

Public Property Get Bits() As Variant
    Bits = tDevMode.lngBits
End Property
Public Property Let Bits(ByVal vNewValue As Variant)
    tDevMode.lngBits = vNewValue
End Property

Public Property Get PW() As Variant
    PW = tDevMode.lngPW
End Property
Public Property Let PW(ByVal vNewValue As Variant)
    tDevMode.lngPW = vNewValue
End Property

Public Property Get PH() As Variant
    PH = tDevMode.lngPH
End Property
Public Property Let PH(ByVal vNewValue As Variant)
    tDevMode.lngPH = vNewValue
End Property

Public Property Get DFI() As Variant
    DFI = tDevMode.lngDFI
End Property
Public Property Let DFI(ByVal vNewValue As Variant)
    tDevMode.lngDFI = vNewValue
End Property

Public Property Get DFr() As Variant
    DFr = tDevMode.lngDFr
End Property
Public Property Let DFr(ByVal vNewValue As Variant)
    tDevMode.lngDFr = vNewValue
End Property
