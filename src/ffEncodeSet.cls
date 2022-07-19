VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffEncodeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' ffmpeg-vba v0.1
'
' An ffmpeg wrapper written in MS Excel VBA (MS Windows)
'
' (c) GCUser99
'
' https://github.com/GCuser99/ffmpeg-vba/tree/main
'
' ==========================================================================
Private mCrf As Integer
Private mPreset As ffPreset
Private mVideoEncoder As ffVideoEncoder
Private mBitRate As String
Private mAdditionalParms As String
Private mAudioEncoder As ffAudioEncoder
Private mPixelFormat As ffPixelFormat

Public Enum ffTransposeCode
    ffRot90CCWFlipVert = 0
    ffRot90CW = 1
    ffRot90CCW = 2
    ffRot90CWFlipVert = 3
    ffRot180CW = 4
    ffFlipLeftRight = 5
End Enum

Public Enum ffMetaRotateCode
    ffRot0 = 0
    ffRot90 = 90
    ffRot180 = 180
    ffRot270 = 270
End Enum

Public Enum ffVideoEncoder
    ffVideoDefault
    ffVideoCopy
    ffH264
    ffH265
    ffLibvpx
    ffLibxvid
    ffMpeg4
End Enum

Public Enum ffPreset
    ffUltrafast
    ffSuperfast
    ffVeryfast
    ffFaster
    ffFast
    ffMedium '– default preset
    ffSlow
    ffSlower
    ffVeryslow
End Enum

Public Enum ffAudioEncoder
    ffAudioDefault
    ffAudioCopy
    ffAAC
    ffMP2
    ffAC3
    ffMWAV2
    ffLibmp3lame
    ffLibvorbis
End Enum

Public Enum ffPixelFormat
    ffPixelDefault
    ffYuv420p
    ffYuvj420p
    ffRGBA
End Enum

Public Property Let BitRate(ByVal val As String)
    mBitRate = val
End Property

Public Property Get BitRate() As String
    BitRate = mBitRate
End Property

Public Property Let AdditionalParms(ByVal val As String)
    'a string containing any other formating options not directly provided in this class
    mAdditionalParms = val
End Property

Public Property Get AdditionalParms() As String
    AdditionalParms = mAdditionalParms
End Property

Public Property Let Crf(ByVal val As Integer)
    mCrf = val
End Property

Public Property Get Crf() As Integer
    Crf = mCrf
End Property

Public Property Let Preset(ByVal val As ffPreset)
    mPreset = val
End Property

Public Property Get Preset() As ffPreset
    Preset = mPreset
End Property

Public Property Let VideoEncoder(ByVal val As ffVideoEncoder)
    mVideoEncoder = val
End Property

Public Property Get VideoEncoder() As ffVideoEncoder
    VideoEncoder = mVideoEncoder
End Property

Public Property Let PixelFormat(ByVal val As ffPixelFormat)
    mPixelFormat = val
End Property

Public Property Get PixelFormat() As ffPixelFormat
    PixelFormat = mPixelFormat
End Property

Public Function GetEncodeString() As String
    Dim encodeStr As String
    encodeStr = ""
    If mCrf >= 0 Then
        encodeStr = encodeStr & " -crf " & mCrf
    End If
    If mVideoEncoder <> ffVideoDefault Then
        encodeStr = encodeStr & " -c:v " & GetVideoEncoderString(mVideoEncoder)
    End If
    If mPreset <> ffMedium Then
        encodeStr = encodeStr & " -preset " & GetPresetString(mPreset)
    End If
    If mBitRate <> "" Then
        encodeStr = encodeStr & " -b:v " & mBitRate
    End If
    If mAudioEncoder <> ffAudioDefault Then
        encodeStr = encodeStr & " -c:a " & GetAudioEncoderString(mAudioEncoder)
    End If
    If mPixelFormat <> ffPixelDefault Then
        encodeStr = encodeStr & " -pix_fmt " & GetPixelFormatString(mPixelFormat)
    End If
    If mAdditionalParms <> "" Then
        encodeStr = encodeStr & " " & VBA.Trim(mAdditionalParms)
    End If
    GetEncodeString = encodeStr
End Function

Private Function GetVideoEncoderString(ByVal encoderVal As ffVideoEncoder) As String
    Select Case encoderVal
    Case ffVideoEncoder.ffVideoDefault
        GetVideoEncoderString = ""
    Case ffVideoEncoder.ffH264
        GetVideoEncoderString = "libx264"
    Case ffVideoEncoder.ffH265
        GetVideoEncoderString = "libx265"
    Case ffVideoEncoder.ffLibvpx
        GetVideoEncoderString = "libvpx-vp9"
    Case ffVideoEncoder.ffVideoCopy
        GetVideoEncoderString = "copy"
    Case ffVideoEncoder.ffLibxvid
        GetVideoEncoderString = "libxvid"
    Case ffVideoEncoder.ffMpeg4
        GetVideoEncoderString = "mpeg4"
    End Select
End Function

Private Function GetPixelFormatString(ByVal val As ffPixelFormat) As String
    Select Case val
    Case ffPixelFormat.ffPixelDefault
        GetPixelFormatString = ""
    Case ffPixelFormat.ffRGBA
        GetPixelFormatString = "rgba"
    Case ffPixelFormat.ffYuv420p
        GetPixelFormatString = "yuv420p"
    Case ffPixelFormat.ffYuvj420p
        GetPixelFormatString = "yuvj420p"
    End Select
End Function

Private Function GetAudioEncoderString(ByVal encoderVal As ffAudioEncoder) As String
    Select Case encoderVal
    Case ffAudioEncoder.ffAudioDefault
        GetAudioEncoderString = ""
    Case ffAudioEncoder.ffAAC
        GetAudioEncoderString = "acc"
    Case ffAudioEncoder.ffAC3
        GetAudioEncoderString = "ac3"
    Case ffAudioEncoder.ffMP2
        GetAudioEncoderString = "mp2"
    Case ffAudioEncoder.ffMWAV2
        GetAudioEncoderString = "mwav2"
    Case ffAudioEncoder.ffLibmp3lame
        GetAudioEncoderString = "libmp3lame"
    Case ffAudioEncoder.ffLibvorbis
        GetAudioEncoderString = "libvorbis"
    Case ffAudioEncoder.ffAudioCopy
        GetAudioEncoderString = "copy"
    End Select
End Function

Private Function GetPresetString(ByVal presetVal As ffPreset) As String
    Select Case presetVal
    Case ffPreset.ffUltrafast
        GetPresetString = "ultrafast"
    Case ffPreset.ffSuperfast
        GetPresetString = "superfast"
    Case ffPreset.ffVeryfast
        GetPresetString = "veryfast"
    Case ffPreset.ffFaster
        GetPresetString = "faster"
    Case ffPreset.ffFast
        GetPresetString = "fast"
    Case ffPreset.ffMedium
        GetPresetString = "medium"
    Case ffPreset.ffSlow
        GetPresetString = "slow"
    Case ffPreset.ffSlower
        GetPresetString = "slower"
    Case ffPreset.ffVeryslow
        GetPresetString = "veryslow"
    End Select
End Function

Private Sub Class_Initialize()
    mCrf = -1
    mPreset = ffMedium
    mVideoEncoder = ffVideoDefault
    mAudioEncoder = ffAudioDefault
End Sub
