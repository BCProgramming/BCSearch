VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CID3TagData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ID3 Tag data...



Private mvarFilename As String

Private mChannelsPerSample As Long, mSamplesPerSecond As Long, mBitsPerSecond As Long, mCopyRighted As Boolean, mDurationSec As Double, mCompression As String
Private mSongTitle As String, mSongArtist As String, mSongAlbum As String, mSongYear As Long, mSongComment As String, mSongTrackNumber As Long, mSongGenre As String, mSongComposer As String, mSongOriginalArtist As String, mSongCopyright As String, mSongURL As String, mSongEncodedBy As String


'NOTE: add Property Get's for above private variables!
Public Function Init(ByVal mfilename As String) As Boolean
    mvarFilename = mfilename
    If modMP3.MP3read_HasTag_ID3v2(mvarFilename) = False Then
        Init = False
    Else
        Init = True
         If modMP3.MP3read_GetAudioStreamInfo(mvarFilename, mChannelsPerSample, mSamplesPerSecond, mBitsPerSecond, mCopyRighted, mDurationSec, mCompression) Then
         
            modMP3.MP3read_GetTagInfo_ID3v2 mvarFilename, mSongTitle, mSongArtist, mSongAlbum, mSongYear, mSongComment, mSongTrackNumber, mSongGenre, mSongComposer, mSongOriginalArtist, mSongCopyright, mSongURL, mSongEncodedBy
         
            
         End If
        
        
        
        
        
    End If
    
    Stop
    
End Function
