VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExifData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfilename As String
Private m_img As GDIPImage
Private Type mExifProperty
    Exif_Name As String
    Exif_Data As Variant
End Type
Private mExif() As mExifProperty

Private mPropertyNames() As String
Private mPropertyValues() As Variant
Public Function GetPropertyNames(ByRef count As Long) As String()


    GetPropertyNames = mPropertyNames
    count = UBound(mPropertyNames)


End Function
Public Function GetPropertyValues(ByRef count As Long) As Variant()
    GetPropertyValues = mPropertyValues
    count = UBound(mPropertyValues)
End Function
Public Function PropValueByName(ByVal Name As String) As Variant

    Dim I As Long
    For I = 1 To UBound(mPropertyNames)
        If StrComp(mPropertyNames(I), Name, vbTextCompare) = 0 Then
            PropValueByName = mPropertyValues(I)
            Exit Function
        End If
    
    Next I


End Function
Public Sub LoadEXIF(ByVal StrFile As String)
    loadImageAndTags StrFile
    Dim I As Long, propcount As Long
    Dim StrBuild As String
    For I = 1 To propcount
    
    
    Next I
End Sub

'Public Sub test()
'loadImageAndTags "D:\Radeon9k.jpg"
'
'
'End Sub
Private Sub GetEXIFTags(img As GDIPImage)
   Dim I As Long
   Dim sItem As String, strfull As String
   Dim prop As GDIPPropertyItem
   Dim j As Long
   Dim v As Variant
   Dim s As String
   Dim Temp As Variant
   ReDim mPropertyNames(1 To img.PropertyCount)
   ReDim mPropertyValues(1 To img.PropertyCount)
   For I = 1 To img.PropertyCount
      Set prop = img.PropertyItem(I)
      'sItem = prop.Name & " "
      mPropertyNames(I) = prop.Name
      Select Case prop.ItemType
      Case PropertyTagTypeASCII
         mPropertyValues(I) = prop.ParseString()
         
      Case PropertyTagTypeRational, PropertyTagTypeSRational
         ReDim Temp(1 To prop.ValueCount)
         For j = 1 To prop.ValueCount
            
            v = prop.ParseRational(j)
            Temp(j) = v(1) / v(2)
         Next j
         mPropertyValues(I) = Temp
      Case PropertyTagTypeLong
         ReDim Temp(1 To prop.ValueCount)
         For j = 1 To prop.ValueCount
            
            v = prop.ParseLong(j)
            Temp(j) = v
         Next j
         mPropertyValues(I) = Temp
      
      Case PropertyTagTypeShort
         ReDim Temp(1 To prop.ValueCount)
         For j = 1 To prop.ValueCount
            
            v = prop.ParseShort(j)
            Temp(j) = v
         Next j
         mPropertyValues(I) = Temp
      
      Case PropertyTagTypeUndefined
         ReDim b(0 To prop.ValueCount - 1) As Byte
         prop.GetData b
         ReDim Temp(0 To prop.ValueCount - 1) As Variant
         For j = 1 To prop.ValueCount - 1
            
            Temp(j) = b(j - 1)
            
            
         Next j
        mPropertyValues(I) = Temp
      Case Else
         mPropertyValues(I) = sItem & prop.ItemType & " " & prop.Length
      End Select
      'lstEXIFTags.AddItem sItem
      
      'strfull = strfull & vbCrLf & sItem
      
   Next I
   'GetEXIFTags = strfull
End Sub

Public Function GetGDIPImage() As Object
    GetGDIPImage = m_img
End Function


Private Sub loadImageAndTags(ByVal sFile As String)

         Dim strshow As String
   On Error GoTo errorHandler
   ' Load the image:
   If Not m_img Is Nothing Then
      m_img.Dispose
      Set m_img = Nothing
   End If
   Set m_img = New GDIPImage
   m_img.FromFile sFile
   
   ' Read EXIFTags:
    GetEXIFTags m_img
   
   'Stop
            
   
   Exit Sub
   
errorHandler:
   'MsgBox "An error occurred trying to show the information for this file: " & Err.Description, vbInformation
   Set m_img = Nothing
   Err.Raise 9
   Exit Sub
   
End Sub

Private Sub Class_Initialize()
GDIPlusCreate
End Sub

Private Sub Class_Terminate()
 If Not m_img Is Nothing Then
      m_img.Dispose
      Set m_img = Nothing
   End If

GDIPlusDispose
End Sub
