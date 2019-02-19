VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrtMip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' [twip]-->[mm]
Private Const C_TWIP_TO_MM = 1000 / 567
' [mm]  -->[twip]
Private Const C_MM_TO_TWIP = 567 / 1000

'Private Const C_HORIZON = 1953
'Private Const C_VERTICAL = 1954

Private Type RawDataType
    RGB As String * 28
End Type

Private Type PrtMipType

    lngLeftMargin As Long
    lngTopMargin As Long
    lngRightMargin As Long
    lngBottomMargin As Long
    lngDataOnly As Long

    lngItemSizeWidth As Long
    lngItemSizeHeight As Long
    lngDefaultSize As Long
    lngColumns As Long
    lngColumnSpacing As Long

    lngRowSpacing As Long
    lngItemLayout As Long
    lngFastPrint As Long
    lngDatasheet As Long

End Type

Dim tPrtMip As PrtMipType

Public Sub LoadData(ByVal vNewValue As Variant)
    Dim rawdata As RawDataType
    rawdata.RGB = vNewValue
    LSet tPrtMip = rawdata
End Sub

Public Function ToString() As String
    Dim rawdata As RawDataType
    LSet rawdata = tPrtMip
    ToString = rawdata.RGB
End Function

Public Property Get LeftMargin_mm() As Variant
    LeftMargin_mm = tPrtMip.lngLeftMargin * C_TWIP_TO_MM
End Property
Public Property Let LeftMargin_mm(ByVal mm As Variant)
    tPrtMip.lngLeftMargin = mm * C_MM_TO_TWIP
End Property

Public Property Get TopMargin_mm() As Variant
    TopMargin_mm = tPrtMip.lngTopMargin * C_TWIP_TO_MM
End Property
Public Property Let TopMargin_mm(ByVal mm As Variant)
    tPrtMip.lngTopMargin = mm * C_MM_TO_TWIP
End Property

Public Property Get RightMargin_mm() As Variant
    RightMargin_mm = tPrtMip.lngRightMargin * C_TWIP_TO_MM
End Property
Public Property Let RightMargin_mm(ByVal mm As Variant)
    tPrtMip.lngRightMargin = mm * C_MM_TO_TWIP
End Property

Public Property Get BottomMargin_mm() As Variant
    BottomMargin_mm = tPrtMip.lngBottomMargin * C_TWIP_TO_MM
End Property
Public Property Let BottomMargin_mm(ByVal mm As Variant)
    tPrtMip.lngBottomMargin = mm * C_MM_TO_TWIP
End Property

Public Property Get ItemSizeWidth_mm() As Variant
    ItemSizeWidth_mm = tPrtMip.lngItemSizeWidth * C_TWIP_TO_MM
End Property
Public Property Let ItemSizeWidth_mm(ByVal mm As Variant)
    tPrtMip.lngItemSizeWidth = mm * C_MM_TO_TWIP
End Property

Public Property Get ItemSizeHeight_mm() As Variant
    ItemSizeHeight_mm = tPrtMip.lngItemSizeHeight * C_TWIP_TO_MM
End Property
Public Property Let ItemSizeHeight_mm(ByVal mm As Variant)
    tPrtMip.lngItemSizeHeight = mm * C_MM_TO_TWIP
End Property

Public Property Get LeftMargin_twip() As Variant
    LeftMargin_twip = tPrtMip.lngLeftMargin
End Property
Public Property Let LeftMargin_twip(ByVal vNewValue As Variant)
    tPrtMip.lngLeftMargin = vNewValue
End Property

Public Property Get TopMargin_twip() As Variant
    TopMargin_twip = tPrtMip.lngTopMargin
End Property
Public Property Let TopMargin_twip(ByVal vNewValue As Variant)
    tPrtMip.lngTopMargin = vNewValue
End Property

Public Property Get RightMargin_twip() As Variant
    RightMargin_twip = tPrtMip.lngRightMargin
End Property
Public Property Let RightMargin_twip(ByVal vNewValue As Variant)
    tPrtMip.lngRightMargin = vNewValue
End Property

Public Property Get BottomMargin_twip() As Variant
    BottomMargin_twip = tPrtMip.lngBottomMargin
End Property
Public Property Let BottomMargin_twip(ByVal vNewValue As Variant)
    tPrtMip.lngBottomMargin = vNewValue
End Property

Public Property Get DataOnly() As Variant
    DataOnly = tPrtMip.lngDataOnly
End Property
Public Property Let DataOnly(ByVal vNewValue As Variant)
    tPrtMip.lngDataOnly = vNewValue
End Property

Public Property Get ItemSizeWidth_twip() As Variant
    ItemSizeWidth_twip = tPrtMip.lngItemSizeWidth
End Property
Public Property Let ItemSizeWidth_twip(ByVal vNewValue As Variant)
    tPrtMip.lngItemSizeWidth = vNewValue
End Property

Public Property Get ItemSizeHeight_twip() As Variant
    ItemSizeHeight_twip = tPrtMip.lngItemSizeHeight
End Property
Public Property Let ItemSizeHeight_twip(ByVal vNewValue As Variant)
    tPrtMip.lngItemSizeHeight = vNewValue
End Property

Public Property Get DefaultSize() As Variant
    DefaultSize = tPrtMip.lngDefaultSize
End Property
Public Property Let DefaultSize(ByVal vNewValue As Variant)
    tPrtMip.lngDefaultSize = vNewValue
End Property

Public Property Get Columns() As Variant
    Columns = tPrtMip.lngColumns
End Property
Public Property Let Columns(ByVal vNewValue As Variant)
    tPrtMip.lngColumns = vNewValue
End Property

Public Property Get ColumnSpacing() As Variant
    ColumnSpacing = tPrtMip.lngColumnSpacing
End Property
Public Property Let ColumnSpacing(ByVal vNewValue As Variant)
    tPrtMip.lngColumnSpacing = vNewValue
End Property

Public Property Get RowSpacing() As Variant
    RowSpacing = tPrtMip.lngRowSpacing
End Property
Public Property Let RowSpacing(ByVal vNewValue As Variant)
    tPrtMip.lngRowSpacing = vNewValue
End Property

Public Property Get ItemLayout() As Variant
    ItemLayout = tPrtMip.lngItemLayout
End Property
Public Property Let ItemLayout(ByVal vNewValue As Variant)
    tPrtMip.lngItemLayout = vNewValue
End Property

Public Property Get FastPrint() As Variant
    FastPrint = tPrtMip.lngFastPrint
End Property
Public Property Let FastPrint(ByVal vNewValue As Variant)
    tPrtMip.lngFastPrint = vNewValue
End Property

Public Property Get Datasheet() As Variant
    Datasheet = tPrtMip.lngDatasheet
End Property
Public Property Let Datasheet(ByVal vNewValue As Variant)
    tPrtMip.lngDatasheet = vNewValue
End Property
