VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' ffmpeg-vba v0.3
'
' An ffmpeg wrapper written in MS Excel VBA (MS Windows)
'
' (c) GCUser99
'
' https://github.com/GCuser99/ffmpeg-vba/tree/main
'
' ==========================================================================
Option Base 1

Private commandWindowStyle As VbAppWinStyle
Private ffplayExecPath As String
Private tmpDirPath As String
Private mDefaultIOPath As String

Private Type ffStreamInfo
    HasVideo As Boolean
    HasAudio As Boolean
    videoCodec As String
    audioCodec As String
    videoDuration As Single
    audioDuration As Single
    PixelFormat As String
    videoWidth As Long
    videoHeight As Long
End Type

Public Property Let DefaultIOPath(ByVal val As String)
    mDefaultIOPath = GetAbsolutePath(val)
End Property

Public Property Get DefaultIOPath() As String
    DefaultIOPath = mDefaultIOPath
End Property

'Note for future: need to use combination of -loop and -autoexit to view image files for specific duration
Public Sub Play(ByVal inputFilepath As String, Optional ByVal startTime As Integer = 0, Optional ByVal endTime As Integer = -1, Optional ByVal autoExit = True, Optional ByVal scaleWindow As Single = 1#, Optional ByVal fullScreen As Boolean = False)

    If startTime > 0 Then cmd = cmd & " -ss " & startTime
    If endTime > 0 Then cmd = cmd & " -t " & (endTime - startTime)
    
    If autoExit Then cmd = cmd & " -autoexit"
    If fullScreen Then
        cmd = cmd & " -fs"
    Else
        If scaleWindow <> 1# Then
            Dim streamInfo As ffStreamInfo
            streamInfo = GetInfo(inputFilepath)
            w = CInt(streamInfo.videoWidth * scaleWindow)
            h = CInt(streamInfo.videoHeight * scaleWindow)
            cmd = cmd & " -x " & w & " -y " & h
        End If
    End If

    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    cmd = cmd & " -i " & instream

    cmd = cmd & " " & VBA.Trim(commandStr)

    'Debug.Print "ffplay.exe " & cmd 'this is used for testing in manually opened command window

    SendCommand cmd
End Sub

Public Sub RunCommand(ByVal commandStr As String, Optional ByVal inputFilepath As String)
    cmd = ""
    If Not IsMissing(inputFilepath) Then
        instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
        cmd = cmd & " -i " & instream & " "
    End If
    
    cmd = cmd & VBA.Trim(commandStr)
    
    'Debug.Print "ffplay.exe " & cmd 'this is used for testing in manually opened command window

    SendCommand cmd
End Sub

Private Function GetAbsolutePath(ByVal strPath As String, Optional ByVal refPath As String = "") As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    If refPath = "" Then refPath = ThisWorkbook.path
    ChDrive refPath
    ChDir refPath
    GetAbsolutePath = fso.GetAbsolutePathName(VBA.Trim(strPath))
    ChDrive saveppath
    ChDir saveppath
End Function

Private Sub SendCommand(ByVal commandStr As String)
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim waitOnReturn As Boolean
    
    waitOnReturn = True
    
    ffplayExecPath = GetAbsolutePath(ffplayExecPath)
    
    ffplayexe = Chr(34) & ffplayExecPath & Chr(34)
    
    If Left(commandStr, 1) <> " " Then commandStr = " " & commandStr
    
    wsh.Run ffplayexe & commandStr, commandWindowStyle, waitOnReturn
End Sub

Private Sub Class_Initialize()
    commandWindowStyle = vbNormalFocus 'vbHide 'hides the command window
    ffplayExecPath = ".\ffplay.exe"
    tmpDirPath = Environ("TEMP") 'same as fso.GetSpecialFolder(2)
    mDefaultIOPath = GetAbsolutePath(".\")
End Sub

Private Function TimeStringToSecs(ByVal ts As String) As Double
    tsparts = Split(ts, ":")
    TimeStringToSecs = 0
    j = 0
    For i = UBound(tsparts) To 0 Step -1
        TimeStringToSecs = TimeStringToSecs + tsparts(i) * (60 ^ j)
        j = j + 1
    Next i
End Function

Private Function GetInfo(ByVal inputFilepath As String) As ffStreamInfo
    Dim jc As New JSonConverter
    Dim prober As New ffProbe
    Dim d As Dictionary
    Dim streamInfo As ffStreamInfo
    Dim streams As Collection

    commandStr = "-hide_banner -print_format json -show_entries stream=codec_name,codec_type,duration,pix_fmt,width,height"
    
    prober.DefaultIOPath = mDefaultIOPath
    resp = prober.RunCommand(commandStr, inputFilepath)
    
    If Not jc.ParseJSON(resp).Exists("streams") Then
        GetInfo = streamInfo
        Exit Function
    End If
    
    Set streams = jc.ParseJSON(resp)("streams")
    
    If streams.Count > 0 Then
        For i = 1 To streams.Count
            If streams(i)("codec_type") = "audio" Then
                streamInfo.HasAudio = True
                streamInfo.audioCodec = streams(i)("codec_name")
                streamInfo.audioDuration = streams(i)("duration")
            End If
            If streams(i)("codec_type") = "video" Then
                streamInfo.HasVideo = True
                streamInfo.videoCodec = streams(i)("codec_name")
                streamInfo.videoDuration = streams(i)("duration")
                streamInfo.PixelFormat = streams(i)("pix_fmt")
                streamInfo.videoWidth = streams(i)("width")
                streamInfo.videoHeight = streams(i)("height")
            End If
        Next i
    End If
    
    GetInfo = streamInfo
End Function
