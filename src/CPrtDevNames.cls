VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrtDevNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Type DevNameOffsetType
    intDriverOffset As Integer
    intDeviceOffset As Integer
    intOutputOffset As Integer
    intDefault As Integer
End Type

Private Type RawDataType
    RGB As String * 256
End Type

Dim tOffset As DevNameOffsetType
Dim rawdata As RawDataType

Dim strDriverName As String
Dim strDeviceName As String
Dim strOutputName As String

Public Sub LoadData(ByVal vNewValue As Variant)
    rawdata.RGB = vNewValue
    LSet tOffset = rawdata
End Sub

Public Function ToString() As String
    Dim typeDevName As DevNameOffsetType
    Dim rdt As RawDataType
    Dim strTemp As String
    
    typeDevName.intDriverOffset = Me.DriverOffset
    
    typeDevName.intDeviceOffset _
        = typeDevName.intDriverOffset + LenB(StrConv(Me.DriverName, vbFromUnicode)) + 1
    
    typeDevName.intOutputOffset _
        = typeDevName.intDeviceOffset + LenB(StrConv(Me.DeviceName, vbFromUnicode)) + 1
    
    typeDevName.intDefault = Me.IsDefault
    
    LSet rdt = typeDevName
    
    strTemp = Me.DriverName & Chr(0) & Me.DeviceName & Chr(0) & Me.OutputName & Chr(0)
            
    ToString = LeftB(rdt.RGB, 8) & StrConv(strTemp, vbFromUnicode)
                        
End Function

Public Property Get DriverOffset() As Variant
    DriverOffset = tOffset.intDriverOffset
End Property
Public Property Let DriverOffset(ByVal vNewValue As Variant)
    tOffset.intDriverOffset = vNewValue
End Property

Public Property Get DeviceOffset() As Variant
    DeviceOffset = tOffset.intDeviceOffset
End Property
Public Property Let DeviceOffset(ByVal vNewValue As Variant)
    tOffset.intDeviceOffset = vNewValue
End Property

Public Property Get OutputOffset() As Variant
    OutputOffset = tOffset.intOutputOffset
End Property
Public Property Let OutputOffset(ByVal vNewValue As Variant)
    tOffset.intOutputOffset = vNewValue
End Property

Public Property Get IsDefault() As Variant
    IsDefault = tOffset.intDefault
End Property
Public Property Let IsDefault(ByVal vNewValue As Variant)
    tOffset.intDefault = vNewValue
End Property

Private Function GetDriverNameFromRawData() As Variant
    Dim idx As Integer
    Dim v As Variant
    v = MidB(rawdata.RGB, Me.DriverOffset + 1)
    idx = InStrB(v, StrConv(Chr(0), vbFromUnicode))
    v = LeftB(v, idx - 1)
    GetDriverNameFromRawData = StrConv(v, vbUnicode)
End Function

Private Function GetDeviceNameFromRawData() As Variant
    Dim idx As Integer
    Dim v As Variant
    idx = InStrB(Me.DeviceOffset + 1, rawdata.RGB, StrConv(Chr(0), vbFromUnicode))
    v = MidB(rawdata.RGB, Me.DeviceOffset + 1, (idx - 1) - Me.DeviceOffset)
    GetDeviceNameFromRawData = StrConv(v, vbUnicode)
End Function

Private Function GetOutputNameFromRawData() As Variant
    Dim idx As Integer
    Dim v As Variant
    idx = InStrB(Me.OutputOffset + 1, rawdata.RGB, StrConv(Chr(0), vbFromUnicode))
    v = MidB(rawdata.RGB, Me.OutputOffset + 1, (idx - 1) - Me.OutputOffset)
    GetOutputNameFromRawData = StrConv(v, vbUnicode)
End Function

Public Property Get DriverName() As String
    If Len(strDriverName) = 0 Then
        strDriverName = GetDriverNameFromRawData()
    End If
    DriverName = strDriverName
End Property
Public Property Let DriverName(ByVal vNewValue As String)
    strDriverName = vNewValue
End Property

Public Property Get DeviceName() As String
    If Len(strDeviceName) = 0 Then
        strDeviceName = GetDeviceNameFromRawData()
    End If
    DeviceName = strDeviceName
End Property
Public Property Let DeviceName(ByVal vNewValue As String)
    strDeviceName = vNewValue
End Property

Public Property Get OutputName() As String
    If Len(strOutputName) = 0 Then
        strOutputName = GetOutputNameFromRawData()
    End If
    OutputName = strOutputName
End Property
Public Property Let OutputName(ByVal vNewValue As String)
    strOutputName = vNewValue
End Property

