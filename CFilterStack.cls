VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IStreamFilter
'CFilterStack: holds a <Group> if IFilterStream objects, and when required delegates all of it's own IFilterStream methods to the collection.


'for example- adding a CCoreFilter set to compress and a CCoreFilter set to encrypt will cause the passed Input stream to be compressed, and then encrypted when passed through.


Private mFilterObjects As Collection 'Collection of IStreamFilter objects.


Public Sub Add(FilterAdd As IStreamFilter, Optional ByVal sKey As String = vbNullString)
    If sKey = vbNullString Then
        mFilterObjects.Add FilterAdd
    Else
        mFilterObjects.Add FilterAdd, sKey
    End If
End Sub

Public Function Item(ByVal vntIndexkey As Variant)
    Set Item = mFilterObjects.Item(vntIndexkey)
End Function
Public Sub Remove(ByVal vntIndexkey As Variant)
    mFilterObjects.Remove vntIndexkey



End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mFilterObjects.[_NewEnum]
End Function
Private Sub Class_Initialize()
    Set mFilterObjects = New Collection
End Sub
Public Sub FilterStream(inputstream As IInputStream, outputstream As IOutputStream, ByVal flReverse As Boolean)
    IStreamFilter_FilterStream inputstream, outputstream, flReverse
End Sub

Private Sub IStreamFilter_FilterStream(inputstream As IInputStream, outputstream As IOutputStream, Optional ByVal flReverse As Boolean = False)

'run through every item in mFilterobjects,taking the output from each and using it as the input for the next one.
Dim idxStart As Integer, IdxEnd As Integer, usestep As Integer
Dim CurrIndex As Integer


If flReverse Then
    idxStart = mFilterObjects.count
    IdxEnd = 1
    usestep = -1
Else
    idxStart = 1
    IdxEnd = mFilterObjects.count
    usestep = 1
End If

Dim currfilter As IStreamFilter


Dim TempOutput As BCFile.IOutputStream
Dim tempinput As BCFile.IInputStream

If mFilterObjects.count = 1 Then
    mFilterObjects.Item(1).FilterStream inputstream, outputstream, flReverse
    Exit Sub
ElseIf mFilterObjects.count = 0 Then

    outputstream.WriteBytes inputstream.readbytes(inputstream.Size)
    Exit Sub
End If

'start off with a memory output stream...

'and use the initial input stream...
Set tempinput = inputstream

For CurrIndex = idxStart To IdxEnd Step usestep
    Set currfilter = mFilterObjects.Item(CurrIndex)
    If CurrIndex = IdxEnd Then
        'for the last item, read from tempinput to our actual output stream...
        currfilter.FilterStream tempinput, outputstream, flReverse
    Else
        'create a new memory stream to hold the output.
        Set TempOutput = New CMemStream
        'read from tempinput into tempoutput...
        currfilter.FilterStream tempinput, TempOutput, flReverse
    
        'And... now make tempinput tempoutput...
        Set tempinput = TempOutput
        tempinput.SeekTo 0, STREAM_BEGIN
        
    
    End If
Next





End Sub
