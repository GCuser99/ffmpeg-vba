VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' ffmpeg-vba v0.2
'
' An ffmpeg wrapper written in MS Excel VBA (MS Windows)
'
' (c) GCUser99
'
' https://github.com/GCuser99/ffmpeg-vba/tree/main
'
' ==========================================================================
Option Base 1

Private ffprobeExecPath As String
Private hideCommandWindow As Boolean

'ShowWindow() Commands
Private Const SW_HIDE = 0
Private Const SW_MINIMIZE = 6
'GetWindow Constants
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Enum ffOutputFormat
    ffDef = 0
    ffJson = 1
    ffCSV = 2
    ffXML = 3
    ffFlat = 4
    ffIni = 5
    ffDict = 6
End Enum

Public Enum ffStreamInfo
    ffAllStreams
    ffFirstVideo
    ffFirstAudio
End Enum

Public Enum ffFileSize
    ffBytes
    ffKiloBytes
    ffMegaBytes
End Enum

Private tmpDirPath As String
Private mDefaultIOPath As String

Public Property Let DefaultIOPath(ByVal val As String)
    mDefaultIOPath = GetAbsolutePath(val)
End Property

Public Property Get DefaultIOPath() As String
    DefaultIOPath = mDefaultIOPath
End Property

Public Function RunCommand(ByVal commandStr As String, Optional ByVal inputFilepath) As String
    If Not IsMissing(inputFilepath) Then
        inputFilepath = Chr(34) & GetAbsolutePath(inputFilepath) & Chr(34)
        cmd = VBA.Trim(commandStr) & " " & inputFilepath
    Else
        cmd = VBA.Trim(commandStr)
    End If
    
    RunCommand = SendCommand(cmd)
End Function

Private Function SendCommand(ByVal commandStr As String) As String
    Dim wsh As New IWshRuntimeLibrary.WshShell, wshExec As IWshRuntimeLibrary.wshExec
    
    ffprobeExecPath = GetAbsolutePath(ffprobeExecPath)
    
    ffprobeexe = Chr(34) & ffprobeExecPath & Chr(34)
    
    If Left(commandStr, 1) <> " " Then commandStr = " " & commandStr
    
    Set wshExec = wsh.Exec(ffprobeexe & commandStr)
    
    If hideCommandWindow Then HideWindow wshExec.ProcessID
    
    'Do While wshExec.Status = 0
    '    DoEvents
    '    Sleep 50
    'Loop
    'Debug.Print wshExec.Status
    
    resp = wshExec.StdOut.ReadAll
    
    SendCommand = resp
End Function

Public Function GetMetaRotation(ByVal inputFilepath As String) As Integer
    'see https://ottverse.com/ffprobe-comprehensive-tutorial-with-examples/
    commandStr = " -loglevel error -select_streams v:0 -show_entries side_data=rotation -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    If resp = "" Then resp = 0
    
    GetMetaRotation = CInt(resp)
End Function

Public Function GetCodecName(ByVal inputFilepath As String) As String
    'see https://ottverse.com/ffprobe-comprehensive-tutorial-with-examples/
    'https://ourcodeworld.com/articles/read/1484/how-to-get-the-information-and-metadata-of-a-media-file-audio-or-video-in-json-format-with-ffprobe
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=codec_name -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetCodecName = resp
End Function

Public Function GetCodecType(ByVal inputFilepath As String) As String
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=codec_type -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetCodecType = resp
End Function

Public Function GetPixelFormat(ByVal inputFilepath As String) As String
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=pix_fmt -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetPixelFormat = resp
End Function

Public Function GetDuration(ByVal inputFilepath As String) As String
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=duration -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetDuration = resp
End Function

Public Function GetFileSize(ByVal inputFilepath As String, Optional sizeUnits As ffFileSize = ffMegaBytes) As String
    commandStr = "-show_entries format=size -v quiet -of csv=""p=0"""
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    Select Case sizeUnits
    Case ffFileSize.ffKiloBytes
        resp = resp / 1000
    Case ffFileSize.ffMegaBytes
        resp = resp / 1000 / 1000
    End Select
    
    GetFileSize = resp
End Function

Public Function GetHeight(ByVal inputFilepath As String) As String
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=height -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetHeight = resp
End Function

Public Function GetWidth(ByVal inputFilepath As String) As String
    commandStr = " -loglevel error -select_streams v:0 -show_entries stream=width -of default=nw=1:nk=1"
    
    resp = Me.RunCommand(commandStr, inputFilepath)
    
    resp = Replace(Trim(resp), vbCrLf, "")
    
    GetWidth = resp
End Function

Public Function GetStreamInfo(ByVal inputFilepath As String, Optional ByVal streamInfo As ffStreamInfo = ffAllStreams, Optional ByVal outputFormat As ffOutputFormat = ffJson) As Variant
    Dim jc As New JSonConverter
    
    fmtStr = GetOutputFormatString(outputFormat)
    
    Select Case streamInfo
    Case ffStreamInfo.ffAllStreams
        cmd = "-v error -hide_banner -print_format " & fmtStr & " -show_streams"
    Case ffStreamInfo.ffFirstVideo
        cmd = "-v error -hide_banner -print_format " & fmtStr & " -select_streams v:0 -show_streams"
    Case ffStreamInfo.ffFirstAudio
        cmd = "-v error -hide_banner -print_format " & fmtStr & " -select_streams a:0 -show_streams"
    End Select
    
    Debug.Print cmd
    
    resp = Me.RunCommand(cmd, inputFilepath)
    
    If outputFormat = ffDict Then
        Set GetStreamInfo = jc.ParseJSON(resp)("streams")
    Else
        GetStreamInfo = resp
    End If
End Function

Public Function HasAudio(ByVal inputFilepath As String) As Boolean
    Dim jc As New JSonConverter
    
    commandStr = "-v error -hide_banner -print_format json -select_streams a:0" & " -show_streams"
    resp = Me.RunCommand(commandStr, inputFilepath)
    If jc.ParseJSON(resp)("streams").Count > 0 Then HasAudio = True Else HasAudio = False
    'If jc.ParseJSON(resp).Exists("streams") Then HasAudio = True Else HasAudio = False
End Function

Public Function HasVideo(ByVal inputFilepath As String) As Boolean
    Dim jc As New JSonConverter
    
    commandStr = "-v error -hide_banner -print_format json -select_streams v:0" & " -show_streams"
    resp = Me.RunCommand(commandStr, inputFilepath)
    If jc.ParseJSON(resp)("streams").Count > 0 Then HasVideo = True Else HasVideo = False
    'If jc.ParseJSON(resp).Exists("streams") Then HasVideo = True Else HasVideo = False
End Function

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

Private Sub Class_Initialize()
    ffprobeExecPath = ".\ffprobe.exe"
    hideCommandWindow = True
    tmpDirPath = Environ("TEMP") 'same as fso.GetSpecialFolder(2)
    mDefaultIOPath = GetAbsolutePath(".\")
End Sub

Private Function HideWindow(iProcessID)
    Dim lngWinHwnd As LongPtr
    Do
        lngWinHwnd = GetHwndFromProcess(CLng(iProcessID))
        DoEvents
    Loop While lngWinHwnd = 0
    Sleep 5 'this is needed before calling showWindow
    HideWindow = ShowWindow(lngWinHwnd, SW_HIDE)
End Function

Private Function GetHwndFromProcess(p_lngProcessId As Long) As LongPtr
    Dim lngDesktop As LongPtr
    Dim lngChild As LongPtr
    Dim lngChildProcessID As Long
    On Error Resume Next
    lngDesktop = GetDesktopWindow()
    lngChild = GetWindow(lngDesktop, GW_CHILD)
    Do While lngChild <> 0
        Call GetWindowThreadProcessId(lngChild, lngChildProcessID)
        If lngChildProcessID = p_lngProcessId Then
            GetHwndFromProcess = lngChild
            Exit Do
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop
    On Error GoTo 0
End Function

Private Function GetOutputFormatString(ByVal outputFormat As ffOutputFormat) As String
    Select Case outputFormat
    Case ffOutputFormat.ffJson, ffOutputFormat.ffDict
        fmtStr = "json"
    Case ffOutputFormat.ffCSV
        fmtStr = "csv"
    Case ffOutputFormat.ffDef
        fmtStr = "default"
    Case ffOutputFormat.ffFlat
        fmtStr = "flat"
    Case ffOutputFormat.ffIni
        fmtStr = "ini"
    Case ffOutputFormat.ffXML
        fmtStr = "xml"
    End Select
    GetOutputFormatString = fmtStr
End Function
