VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffMPeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' ffmpeg-vba v0.4
'
' An ffmpeg wrapper written in MS Excel VBA (MS Windows)
'
' (c) GCUser99
'
' https://github.com/GCuser99/ffmpeg-vba/tree/main
'
' ==========================================================================
'
' INSTALLATION:
'
' First download ffmpeg from:
'
'   https://github.com/BtbN/FFmpeg-Builds/releases
'
'   or...
'
'   https://www.gyan.dev/ffmpeg/builds/
'
' Locate the ffmpeg, ffplay, and ffprobe executables in the same directory as this Excel file.
'
' If you wish to store the executables above in a different directory, then modify the executable paths
' in the respective "Class_Initialize" subs of the FFmpeg, FFplay, and FFprobe classes provided here.
'
' ffmpeg-vba supports relative file paths. The default i/o path directory is same as where this Excel file resides.
'
' However, that can be changed by using the following command once before file specification:
'
'   ffmpeg.DefaultIOPath=[path to your media files]
'
' To specify a file located in the default i/o directory, use either ".\my_file.mp4" or "my_file.mp4"
'
' Some method inputs can take an array of file paths - these arrays can take the following forms:
'
'    - a comma-separated string, for example "file1.mp4, file2.mp4, file3.mp4"
'    - a 1D array, for example Array("file1.mp4", "file2.mp4", "file3.mp4")
'
' The DeleteFiles method does recognize wildcards such as "*".
'
' Parameters that take times can be specified in fractional seconds (eg, 25.5) or as a time stamp "[hr]:[min]:[secs]",
' for example "0:0:25.50".
'
' Some method inputs take an array of time pairs - these can take the following forms:
'
'   - a comma-separated string of time pairs, for example: "0.0 to 5.0, 00:10.0 to 00:15.0, 00:00:20.0 to 00:00:25.0"
'   - a 1D array of time pair strings, for example: Array("0.0 to 5.0", "00:10.0 to 00:15.0", "00:00:20.0 to 00:00:25.0")
'   - or a 2D array of times, for example: Array(Array(0.0,5.0), Array("00:10.0","00:15.0"), Array("00:00:20.0","00:00:25.0"))
'
' ==========================================================================
Option Base 1

Private commandWindowStyle As VbAppWinStyle
Private ffmpegExecPath As String
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

Private Sub SendCommand(ByVal commandStr As String)
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim waitOnReturn As Boolean
    
    waitOnReturn = True
    
    ffmpegExecPath = GetAbsolutePath(ffmpegExecPath)
    
    ffmpegexe = Chr(34) & ffmpegExecPath & Chr(34)
    
    If Left(commandStr, 1) <> " " Then commandStr = " " & commandStr
    
    wsh.Run ffmpegexe & commandStr, commandWindowStyle, waitOnReturn
End Sub

Public Property Let DefaultIOPath(ByVal val As String)
    mDefaultIOPath = GetAbsolutePath(val)
End Property

Public Property Get DefaultIOPath() As String
    DefaultIOPath = mDefaultIOPath
End Property

Public Sub RunCommand(ByVal commandStr As String, Optional ByVal inputFilepathArray As Variant, Optional ByVal outputFilepath As Variant)
    'this will process (multiple) input files if they are supplied, as well as a single output file
    'if not supplied then assumes user encoded inputs and outputs into the command string
    'if user specifies input file paths, then they are specified before the command string (does not handle output seeking)
    
    If Not IsMissing(inputFilepathArray) Then
        If Not IsArray(inputFilepathArray) Then 'maybe comma separated string list...
            inputFilepathArray = Split(inputFilepathArray, ",")
        End If
        
        lb = LBound(inputFilepathArray)
        ub = UBound(inputFilepathArray)
        
        For i = lb To ub
            instream = Chr(34) & GetAbsolutePath(inputFilepathArray(i), mDefaultIOPath) & Chr(34)
            cmd = cmd & " -i " & instream
        Next i
        
        cmd = cmd & " " & VBA.Trim(commandStr)
    Else
        cmd = VBA.Trim(commandStr)
    End If
    
    If Not IsMissing(outputFilepath) Then
        outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
        cmd = cmd & " " & outstream
    End If
    
    'Debug.Print "ffmpeg.exe " & cmd 'this is used for testing in manually opened command window

    SendCommand cmd
End Sub

Sub RemoveVideo(ByVal inputFilepath As String, ByVal outputFilepath As String)
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)

    cmd = " -i " & instream & " -vn -c:a copy -y " & outstream
    
    SendCommand cmd
End Sub

Sub RemoveAudio(ByVal inputFilepath As String, ByVal outputFilepath As String)
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = " -i " & instream & " -c:v copy -an -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub Trim(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional ByVal startTime As String, Optional ByVal endTime As String, Optional ByVal ReEncode As Boolean = False, Optional ByVal useOutputSeeking As Boolean = False, Optional oEncSet As ffEncodeSet)
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    'see https://ottverse.com/trim-cut-video-using-start-endtime-reencoding-ffmpeg/
    
    If IsNumeric(startTime) Then
        st = startTime
    Else
        st = TimeStringToSecs(startTime)
    End If
    
    If IsNumeric(endTime) Then
        et = endTime
    Else
        et = TimeStringToSecs(endTime)
    End If
    
    If ReEncode Then
        'encode to trim accurately
        If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString

        If useOutputSeeking Then
            cmd = " -ss " & st & " -i " & instream & " -to " & et & encodeStr & " -y " & outstream
        Else
            cmd = " -i " & instream & " -ss " & st & " -to " & et & encodeStr & " -y " & outstream
        End If
    Else
        If useOutputSeeking Then
            'this is faster for a large file but not as accurate
            cmd = " -ss " & st & " -i " & instream & " -to " & et & " -c copy -y " & outstream
        Else
            cmd = " -i " & instream & " -ss " & st & " -to " & et & " -c copy -y " & outstream
        End If
    End If
    'Debug.Print cmd
    SendCommand cmd
End Sub

Public Sub ReEncode(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional ByVal useOutputSeeking As Boolean = False, Optional oEncSet As ffEncodeSet)
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    cmd = " -i " & instream & encodeStr & " -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub Copy(ByVal inputFilepath As String, ByVal outputFilepath As String)
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = " -i " & instream & " -c copy -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub Join(ByVal inputFilepathArray As Variant, ByVal outputFilepath As String, Optional ByVal ReEncode As Boolean = True, Optional oEncSet As ffEncodeSet)
    Dim fso As New FileSystemObject
    Dim fileStream As TextStream
    Dim streamInfo As ffStreamInfo
    
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If Not IsArray(inputFilepathArray) Then 'maybe comma separated string list...
        inputFilepathArray = Split(inputFilepathArray, ",")
    End If
    
    lb = LBound(inputFilepathArray)
    ub = UBound(inputFilepathArray)
    
    If ReEncode Then
        'if user didn't specify a parameter then we let ffmpeg use its own defaults
        
        If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
        
        cmd = ""
        cmd2 = ""
        
        'ffmpeg concat assumes that all inputs have same num of channels
        'so this will fail otherwise
        streamInfo = GetInfo(inputFilepathArray(lb))
        
        j = 0
        For i = lb To ub
            instream = Chr(34) & GetAbsolutePath(inputFilepathArray(i), mDefaultIOPath) & Chr(34)
            cmd = cmd & " -i " & instream
            If streamInfo.HasVideo Then cmd2 = cmd2 & "[" & j & ":v]"
            If streamInfo.HasAudio Then cmd2 = cmd2 & "[" & j & ":a]"
            j = j + 1
        Next i
        
        cmd = cmd & " -filter_complex """
        
        cmd = cmd & cmd2 & "concat=n=" & j
        If streamInfo.HasVideo Then cmd = cmd & ":v=1" Else cmd = cmd & ":v=0"
        If streamInfo.HasAudio Then cmd = cmd & ":a=1" Else cmd = cmd & ":a=0"
        If streamInfo.HasVideo Then cmd = cmd & "[outv]"
        If streamInfo.HasAudio Then cmd = cmd & "[outa]"
        cmd = cmd & """"
        If streamInfo.HasVideo Then cmd = cmd & " -map ""[outv]"""
        If streamInfo.HasAudio Then cmd = cmd & " -map ""[outa]"""
        cmd = cmd & encodeStr & " -y " & outstream
        
    Else
        'using the concat demuxer - this assumes that all inputs share same encoding and is VERY fast
        'if spaces in file path, then must be bracketed by Chr(39), not Chr(34)
        filelist = ""
        For i = lb To ub
            filelist = filelist & "file " & Chr(39) & GetAbsolutePath(inputFilepathArray(i), mDefaultIOPath) & Chr(39) & vbCrLf
        Next i
        
        'write the file list to a temp text file
        tmpFilePath = tmpDirPath & "\tmp.txt"
    
        Set fileStream = fso.CreateTextFile(tmpFilePath)
    
        fileStream.Write filelist
        fileStream.Close
        
        cmd = " -f concat -safe 0 -i " & Chr(34) & tmpFilePath & Chr(34) & " -c copy -y " & outstream

    End If
    
    'Debug.Print cmd
    
    SendCommand cmd
    
    If Not ReEncode Then fso.DeleteFile tmpFilePath, True
    
End Sub

Public Sub Edit(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal keepTimesArray As Variant, Optional ByVal ReEncode As Boolean = True, Optional oEncSet As ffEncodeSet)
    Dim fso As New FileSystemObject
    Dim fileStream As TextStream
    Dim streamInfo As ffStreamInfo

    'Edit method can either take a 1d or 2d array or a string like this:
    '0.0 to 5.0, 10 to 20, 0:0:25 to 0:0:30
    'time pairs are comma separated, using the format [starttime1] to [endtime1],[starttime2] to [endtime2] ...
    'if 1d array, then array of strings like this: Array("0.0 to 5.0", "10 to 20", "0:0:25 to 0:0:30")
    
    If Not IsArray(keepTimesArray) Then 'maybe comma separated string list...
        tPairs = Split(keepTimesArray, ",")
        
        ReDim tmpArray(LBound(tPairs) To UBound(tPairs), 1 To 2)
        
        For i = LBound(tPairs) To UBound(tPairs)
            tmpArray(i, 1) = VBA.Trim(Split(tPairs(i), "to", , vbTextCompare)(0))
            tmpArray(i, 2) = VBA.Trim(Split(tPairs(i), "to", , vbTextCompare)(1))
        Next i
        
        keepTimesArray = tmpArray
    Else
        If Not is2d(keepTimesArray) Then
            ReDim tmpArray(LBound(keepTimesArray) To UBound(keepTimesArray), 1 To 2)
            
            For i = LBound(keepTimesArray) To UBound(keepTimesArray)
                tmpArray(i, 1) = VBA.Trim(Split(keepTimesArray(i), "to", , vbTextCompare)(0))
                tmpArray(i, 2) = VBA.Trim(Split(keepTimesArray(i), "to", , vbTextCompare)(1))
            Next i

            keepTimesArray = tmpArray
        End If
    End If
    
    'if not keep pairs but instead reject pairs, could convert to keep pairs here
    'but might need to know start and end times of input to do that?

    If ReEncode Then
        'if user didn't specify a parameter then we let ffmpeg use its own defaults
        
        instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
        outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
        
        streamInfo = GetInfo(inputFilepath)
        
        If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString

        cmd = " -i " & instream & " -filter_complex """
        cmd2 = ""
        
        'look into 2-pass quality-file size
        lb = LBound(keepTimesArray)
        ub = UBound(keepTimesArray)
        
        j = 0
        For i = lb To ub
            If IsNumeric(keepTimesArray(i, 1)) Then
                s = keepTimesArray(i, 1)
            Else
                s = TimeStringToSecs(keepTimesArray(i, 1))
            End If
            
            If IsNumeric(keepTimesArray(i, 2)) Then
                e = keepTimesArray(i, 2)
            Else
                e = TimeStringToSecs(keepTimesArray(i, 2))
            End If
            
            If streamInfo.HasVideo Then cmd = cmd & "[0:v]trim=start=" & s & ":end=" & e & ",setpts=PTS-STARTPTS[" & j & "v];"
            If streamInfo.HasAudio Then cmd = cmd & "[0:a]atrim=start=" & s & ":end=" & e & ",asetpts=PTS-STARTPTS[" & j & "a];"
            If streamInfo.HasVideo Then cmd2 = cmd2 & "[" & j & "v]"
            If streamInfo.HasAudio Then cmd2 = cmd2 & "[" & j & "a]"
            j = j + 1
        Next i
        
        cmd = cmd & cmd2 & "concat=n=" & j
        If streamInfo.HasVideo Then cmd = cmd & ":v=1" Else cmd = cmd & ":v=0"
        If streamInfo.HasAudio Then cmd = cmd & ":a=1" Else cmd = cmd & ":a=0"
        If streamInfo.HasVideo Then cmd = cmd & "[outv]"
        If streamInfo.HasAudio Then cmd = cmd & "[outa]"
        cmd = cmd & """"
        If streamInfo.HasVideo Then cmd = cmd & " -map ""[outv]"""
        If streamInfo.HasAudio Then cmd = cmd & " -map ""[outa]"""
        cmd = cmd & encodeStr & " -y " & outstream
        'Debug.Print cmd
        SendCommand cmd
    Else
        'do Trim n times and then call Join to stitch back together
        ext = fso.GetExtensionName(inputFilepath)
        
        lb = LBound(keepTimesArray)
        ub = UBound(keepTimesArray)
        
        ReDim filelist(0 To ub - lb)
        
        j = 0
        For i = lb To ub
            If IsNumeric(keepTimesArray(i, 1)) Then
                s = keepTimesArray(i, 1)
            Else
                s = TimeStringToSecs(keepTimesArray(i, 1))
            End If
            
            If IsNumeric(keepTimesArray(i, 2)) Then
                e = keepTimesArray(i, 2)
            Else
                e = TimeStringToSecs(keepTimesArray(i, 2))
            End If
            
            tmpInFile = tmpDirPath & "\tmp" & j & "." & ext
            
            Me.Trim inputFilepath, tmpInFile, s, e, False, False
            
            filelist(j) = tmpInFile
            
            j = j + 1
        Next i
        
        Me.Join filelist, outputFilepath, False
        
    End If
End Sub

Public Sub Transpose(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal tcode As ffTransposeCode, Optional oEncSet As ffEncodeSet)
    '0 - rotate by 90 deg counter-clockwise and flip vertically (default)
    '1 - rotate by 90 deg clockwise
    '2 - rotate by 90 deg counter-clockwise
    '3 - rotate by 90 deg clockwise and flip vertically
    '4 - rotate by 90 deg clockwise twice (180 deg)
    '5 - rotate by 90 deg clockwise and flip vertically, then rotate by 90 deg counter-clockwise (flip left-to-right)
    Dim streamInfo As ffStreamInfo

    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    streamInfo = GetInfo(inputFilepath)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
        
    cmd = " -i " & instream & " -filter_complex """
    
    Select Case tcode
    Case 0, 1, 2, 3
        cmd = cmd & "[0:v]transpose=" & tcode & "[outv];"
    Case 4 '180 deg
        cmd = cmd & "[0:v]transpose=1,transpose=1[outv];"
    Case 5 'flip left-right
        cmd = cmd & "[0:v]transpose=3,transpose=2[outv]"
    End Select
    
    If streamInfo.HasAudio Then
        cmd = cmd & ";"
        cmd = cmd & "[0:a]acopy[outa]"
        cmd = cmd & """ -map ""[outv]"" -map ""[outa]""" & encodeStr & " -y " & outstream 'specify outputs and encoding parameters
    Else
        cmd = cmd & """ -map ""[outv]""" & encodeStr & " -y " & outstream 'specify outputs and encoding parameters
    End If
    
    SendCommand cmd
End Sub

Public Sub AutoLevel(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional oEncSet As ffEncodeSet)
    'stretches luminence to full range
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
        
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
        
    cmd = " -i " & instream & " -vf ""pp='al'""" & encodeStr & " -y " & outstream

    SendCommand cmd
End Sub

Public Sub MetaRotate(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal rot As ffMetaRotateCode)
    'this function changes the rotation in metadata without re-encoding
    'note that if there is already a rotation in metadata, this will overwrite it - not add to it
    'must be either 0,90,180,270
    'can use ffmpeg.GetMetaRotation to find existing rotation
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = " -i " & instream & " -map_metadata 0 -metadata:s:v rotate=""" & rot & """ -c copy -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub ExtractFrame(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal frameTime As String, Optional ByVal useOutputSeeking As Boolean = True)
    'frameTime can be specified two ways - either a string hh:mm:ss.000 or a number representing secs such as 10.2
    'ffmpeg -i input.mp4 -ss 00:00:10.000 -vframes 1 thumb.jpg
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If useOutputSeeking Then
        cmd = " -ss " & frameTime & " -i " & instream & " -vframes 1 -y " & outstream
    Else
        cmd = " -i " & instream & " -ss " & frameTime & " -vframes 1 -y " & outstream
    End If
    
    SendCommand cmd
End Sub

Public Sub Reverse(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional oEncSet As ffEncodeSet)
    Dim streamInfo As ffStreamInfo
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    streamInfo = GetInfo(inputFilepath)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    cmd = " -i " & instream
    If streamInfo.HasVideo Then cmd = cmd & " -vf reverse"
    If streamInfo.HasAudio Then cmd = cmd & " -af areverse"
    cmd = cmd & encodeStr & " -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub ChangeSpeed(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal speedFactor As Single, Optional oEncSet As ffEncodeSet)
    Dim streamInfo As ffStreamInfo
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    streamInfo = GetInfo(inputFilepath)
    
    slowfactor = 1 / speedFactor
    
    If streamInfo.HasAudio Then atempo = Application.min(Application.Max(speedFactor, 0.5), 100) 'atempo needs to be between .5 to 100
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    cmd = " -i " & instream & " -filter_complex """
    
    If streamInfo.HasVideo Then cmd = cmd & "[0:v]setpts=" & Format(slowfactor, "0.00") & "*PTS,fifo[outv]"
    If streamInfo.HasVideo And streamInfo.HasAudio Then cmd = cmd & ";"
    If streamInfo.HasAudio Then cmd = cmd & "[0:a]atempo=" & atempo & ",afifo[outa]"
    
    cmd = cmd & """"
    If streamInfo.HasVideo Then cmd = cmd & " -map ""[outv]"""
    If streamInfo.HasAudio Then cmd = cmd & " -map ""[outa]"""
    cmd = cmd & encodeStr & " -y " & outstream 'specify outputs and encoding parameters
    
    SendCommand cmd
End Sub

Sub ChangeVolume(ByVal inputFilepath As String, ByVal outputFilepath As String, ByVal volumeFactor As String)
    'could add time pairs too
    'ffmpeg -i in.mp4 -vcodec copy -af "volume=enable='between(t,5,10)':volume=0, volume=enable='between(t,15,20)':volume=2.0" out.mp4
    'volumeFactor can be a scalar or decibel like "6db"
    Dim streamInfo As ffStreamInfo
    
    streamInfo = GetInfo(inputFilepath)
    
    If Not streamInfo.HasAudio Then
        MsgBox "Must have audio channel to apply volume change!"
        Exit Sub
    End If
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    volumeFactor = Replace(Replace(volumeFactor, "D", "d"), "b", "B")
    
    cmd = " -i " & instream & " -af volume=" & volumeFactor
    
    If streamInfo.HasVideo Then cmd = cmd & " -vcodec copy"

    cmd = cmd & " -y " & outstream

    SendCommand cmd
End Sub

Sub DubOver(ByVal videoInputFilePath As String, ByVal audioInputFilePath As String, ByVal outputFilepath As String, Optional ByVal volumeFactor As String)
    Dim streamInfo As ffStreamInfo
    
    streamInfo = GetInfo(audioInputFilePath)
    
    If Not streamInfo.HasAudio Then
        MsgBox "Input must have audio channel to merge with video!"
        Exit Sub
    End If
    
    instream1 = Chr(34) & GetAbsolutePath(videoInputFilePath, mDefaultIOPath) & Chr(34)
    instream2 = Chr(34) & GetAbsolutePath(audioInputFilePath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If volumeFactor <> "" Then
        volumeFactor = Replace(Replace(volumeFactor, "D", "d"), "b", "B")
        cmd = " -i " & instream1 & " -i " & instream2 & " -c:v copy -map 0:v:0 -map 1:a:0 -shortest -af volume=" & volumeFactor & " -y " & outstream
    Else
        cmd = " -i " & instream1 & " -i " & instream2 & " -c:v copy -map 0:v:0 -map 1:a:0 -shortest -y " & outstream
    End If

    SendCommand cmd
End Sub

Sub Silence(ByVal videoInputFilePath As String, ByVal outputFilepath As String)
    'this zero's the volume of audio track without removing it
    'if there no audio track exists, then it adds a silent one
    Dim streamInfo As ffStreamInfo
    
    streamInfo = GetInfo(videoInputFilePath)

    instream = Chr(34) & GetAbsolutePath(videoInputFilePath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If streamInfo.HasAudio Then
        'could also use ChangeVolume videoInputFilePath, outputFilepath, 0.0
        cmd = " -f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 -i " & instream & " -c:v copy -c:a aac -map 0:a -map 1:v -shortest -y " & outstream
    Else
        cmd = " -f lavfi -i anullsrc=channel_layout=stereo:sample_rate=44100 -i " & instream & " -c:v copy -c:a aac -shortest -y " & outstream
    End If

    SendCommand cmd
End Sub

Public Sub Fade(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional ByVal FadeInDuration As Single = 3, Optional ByVal FadeOutDuration As Single = 3, Optional ByVal fadeAudio As Boolean = True, Optional oEncSet As ffEncodeSet)
    Dim streamInfo As ffStreamInfo
    Dim ffMpeg As New ffMpeg
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    streamInfo = GetInfo(inputFilepath)
    
    outstarttime = streamInfo.videoDuration - FadeOutDuration
    
    If fadeAudio Then
        cmd = " -i " & instream
        If FadeInDuration > 0 Or FadeOutDuration > 0 Then
            cmd = cmd & " -vf """
            If FadeInDuration > 0 Then cmd = cmd & "fade=t=in:st=0:d=" & FadeInDuration & ":color=black"
            If FadeInDuration > 0 And FadeOutDuration > 0 Then cmd = cmd & ","
            If FadeOutDuration > 0 Then cmd = cmd & "fade=t=out:st=" & outstarttime & ":d=" & FadeOutDuration & ":color=black"
            cmd = cmd & """"
        End If
        If FadeInDuration > 0 Or FadeOutDuration > 0 Then
            cmd = cmd & " -af """
            If FadeInDuration > 0 Then cmd = cmd & "afade=t=in:st=0:d=" & FadeInDuration
            If FadeInDuration > 0 And FadeOutDuration > 0 Then cmd = cmd & ","
            If FadeOutDuration > 0 Then cmd = cmd & "afade=t=out:st=" & outstarttime & ":d=" & FadeOutDuration
            cmd = cmd & """"
        End If
        cmd = cmd & encodeStr & " -y " & outstream
    Else
        cmd = " -i " & instream
        If FadeInDuration > 0 Or FadeOutDuration > 0 Then
            cmd = cmd & " -vf """
            If FadeInDuration > 0 Then cmd = cmd & "fade=t=in:st=0:d=" & FadeInDuration & ":color=black"
            If FadeInDuration > 0 And FadeOutDuration > 0 Then cmd = cmd & ","
            If FadeOutDuration > 0 Then cmd = cmd & "fade=t=out:st=" & outstarttime & ":d=" & FadeOutDuration & ":color=black"
            cmd = cmd & """"
        End If
        cmd = cmd & " -c:a copy " & encodeStr & " -y " & outstream
    End If
    
    'Debug.Print cmd
    SendCommand cmd
End Sub

Public Sub CurvesFilter(ByVal inputFilepath As String, ByVal outputFilepath As String, oCurves As ffCurvesFilter, Optional oEncSet As ffEncodeSet)
    Dim ffMpeg As New ffMpeg
    Dim streamInfo As ffStreamInfo
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    streamInfo = GetInfo(inputFilepath)
    'output of curves filter defaults to rgb - here we get the pixel format of the input and use that for output
    PixelFormat = streamInfo.PixelFormat
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    curvesStr = GetFilterString(oCurves) & ",format=" & PixelFormat
    
    cmd = "-i " & instream & " -vf " & """" & curvesStr & """" & encodeStr & " -c:a copy -y " & outstream
    
    SendCommand cmd
End Sub

Public Sub Resize(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional ByVal width As Integer = -1, Optional ByVal height As Integer = -1, Optional ByVal scaleBy As Single = 1#, Optional oEncSet As ffEncodeSet)
    'https://trac.ffmpeg.org/wiki/Scaling
    'if width or height = -1, then keeps original size, scaled by scaleBy
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    If scaleBy <> 1 Then
        scaleStr = "scale=iw*" & scaleBy & ":ih*" & scaleBy
    Else
        If width > 0 And height = -1 Then
            scaleStr = "scale=" & width & ":-2"
        End If
        If height > 0 And width = -1 Then
            scaleStr = "scale=-2:" & height
        End If
        If height > 0 And width > 0 Then
            scaleStr = "scale=" & width & ":" & height
        End If
    End If
    
    cmd = "-i " & instream & " -vf " & scaleStr & encodeStr & " -c:a copy -y " & outstream
    
    'Debug.Print cmd
    SendCommand cmd
End Sub

Public Sub Crop(ByVal inputFilepath As String, ByVal outputFilepath As String, Optional ByVal x As String, Optional ByVal y As String, Optional ByVal width As String, Optional ByVal height As String, Optional oEncSet As ffEncodeSet)
    Dim w As String, h As String

    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    If width = "" Then width = "iw"
    If height = "" Then height = "ih"
    If x = "" Then x = "(iw-out_w)/2"
    If y = "" Then y = "(ih-out_h)/2"
    
    cropStr = """crop=w=" & width & ":h=" & height & ":x=" & x & ":y=" & y & """"
    
    cmd = "-i " & instream & " -vf " & cropStr & encodeStr & " -c:a copy -y " & outstream
    
    'Debug.Print cmd
    SendCommand cmd
End Sub

Public Sub Overlay(ByVal inputFilepath As String, ByVal outputFilepath As String, oOverlays As ffOverlays, Optional oEncSet As ffEncodeSet)
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = " -i " & instream
    ovrly = ""
    For i = 1 To oOverlays.Count
        inp = Chr(34) & GetAbsolutePath(oOverlays.Item(i).InputPath, mDefaultIOPath) & Chr(34)
        st = oOverlays.Item(i).startTime
        et = oOverlays.Item(i).endTime
        fi = oOverlays.Item(i).FadeInDuration
        fo = oOverlays.Item(i).FadeOutDuration
        rs = oOverlays.Item(i).Resize
        xl = oOverlays.Item(i).XLoc
        yl = oOverlays.Item(i).YLoc

        If IsImage(oOverlays.Item(i).InputPath) Then
            sinp = sinp & " -loop 1 -t " & et - st & " -i " & inp
        Else
            sinp = sinp & " -i " & inp
        End If
        
        If rs = 1 Then
            scaleStr = scaleStr & "[" & i & "]copy[rsovr" & i & "];"
        Else
            scaleStr = scaleStr & "[" & i & "]scale=iw*" & rs & ":ih*" & rs & "[rsovr" & i & "];"
        End If
        
        filt = filt & "[rsovr" & i & "]fade=st=0:d=" & fi & ":alpha=1,fade=out:st=" & et - st - fo & ":d=" & fo & ":alpha=1,setpts=PTS+" & st & "/TB[ovr" & i & "];"
        If i = 1 Then
            ovrly = ovrly & "[0:v][ovr" & i & "]overlay=x=" & xl & ":y=" & yl & ":enable='between(t," & st & "," & et & ")'[base" & i & "]"
        Else
            ovrly = ovrly & ";[base" & i - 1 & "][ovr" & i & "]overlay=x=" & xl & ":y=" & yl & ":enable='between(t," & st & "," & et & ")'[base" & i & "]"
        End If
    Next i
    
    cmd = cmd & sinp
    cmd = cmd & " -filter_complex """
    cmd = cmd & scaleStr & filt & ovrly
    cmd = cmd & """ -map """ & "[base" & oOverlays.Count & "]" & """"
    cmd = cmd & " -c:a copy -shortest" & encodeStr & " -y " & outstream
    
    'Debug.Print cmd
    
    SendCommand cmd
End Sub

Public Sub MakeSlideShow(oSlides As ffSlideShow, ByVal outputFilepath As String, Optional oEncSet As ffEncodeSet)
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
    
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = ""
    off = 0
    For i = 1 To oSlides.Count
        inp = Chr(34) & GetAbsolutePath(oSlides(i).InputPath, mDefaultIOPath) & Chr(34)
        d = oSlides(i).Duration
        fd = oSlides(i).TransitionDuration
        tr = GetXfadeTransString(oSlides(i).TransitionType)
        
        If IsImage(oSlides.Item(i).InputPath) Then
            cmd = cmd & " -loop 1 -t " & d & " -i " & inp
        Else
            'cmd = cmd & " -i " & inp
            cmd = cmd & " -t " & d & " -i " & inp
        End If
        
        off = off + d - fd
        If i = 1 Then
            fltStr = "[0][1]xfade=transition=" & tr & ":duration=" & fd & ":offset=" & off & "[f1]"
        ElseIf i < oSlides.Count Then
            fltStr = fltStr & ";[f" & i - 1 & "][" & i & "]xfade=transition=" & tr & ":duration=" & fd & ":offset=" & off & "[f" & i & "]"
        End If
    Next i
    
    cmd = cmd & " -filter_complex " & """" & fltStr & """" & " -map ""[f" & i - 2 & "]"" -r 25 -pix_fmt yuv420p " & encodeStr & " -y "
    
    cmd = cmd & outstream
    
    Debug.Print Replace(Replace(cmd, ";", ";" & vbCr), "-filter_complex """, "-filter_complex """ & vbCr)
    
    SendCommand cmd

End Sub

Public Sub Play(ByVal inputFilepath As String, Optional ByVal startTime As Integer = 0, Optional ByVal endTime As Integer = -1, Optional ByVal autoExit = True, Optional ByVal scaleWindow As Single = 1#, Optional ByVal fullScreen As Boolean = False)
    Dim player As New ffPlay
    player.DefaultIOPath = mDefaultIOPath
    player.Play inputFilepath, startTime, endTime, autoExit, scaleWindow, fullScreen
End Sub

Public Function GetMetaRotation(ByVal inputFilepath As String) As Integer
    GetMetaRotation = Me.Probe.GetMetaRotation(inputFilepath)
End Function

Public Sub DeleteFiles(ParamArray filelist() As Variant)
    Dim fso As New FileSystemObject
    For i = 0 To UBound(filelist)
        If IsArray(filelist(i)) Then
            For j = LBound(filelist(i)) To UBound(filelist(i))
                strlist = Split(filelist(i)(j), ",")
                For K = 0 To UBound(strlist)
                    On Error Resume Next
                    fso.DeleteFile GetAbsolutePath(strlist(K), mDefaultIOPath), True
                    On Error GoTo 0
                Next K
            Next j
        Else
            strlist = Split(filelist(i), ",")
            For K = 0 To UBound(strlist)
                On Error Resume Next
                fso.DeleteFile GetAbsolutePath(strlist(K), mDefaultIOPath), True
                On Error GoTo 0
            Next K
        End If
    Next i
End Sub

Public Sub DrawText(ByVal inputFilepath As String, ByVal outputFilepath As String, oTexts As ffTexts, Optional oEncSet As ffEncodeSet)
    Dim oTextDraw As ffText
    
    If oEncSet Is Nothing Then encodeStr = "" Else encodeStr = oEncSet.GetEncodeString
         
    instream = Chr(34) & GetAbsolutePath(inputFilepath, mDefaultIOPath) & Chr(34)
    outstream = Chr(34) & GetAbsolutePath(outputFilepath, mDefaultIOPath) & Chr(34)
    
    cmd = " -i " & instream & " -vf " & """"

    For i = 1 To oTexts.Count
    
        Set oTextDraw = oTexts(i)
        
        st = oTextDraw.startTime
        et = oTextDraw.endTime
        
        If i > 1 Then cmd = cmd & ","
        
        txt = EscapeSpecialChars(oTextDraw.Text)

        If et = -1 Then
            et = Me.Probe.GetDuration(inputFilepath)
        End If

        cmd = cmd & "drawtext=enable='between(t," & st & "," & et & ")':text='" & txt & "'"
        
        fi = oTextDraw.FadeInDuration
        fo = oTextDraw.FadeOutDuration
        al = oTextDraw.Alpha
        
        If fi + fo > et - st Then 'scale fade durations back to fit within text display time window
            rfact = (et - st) / (fi + fo)
            fi = fi * rfact
            fo = fo * rfact
        End If
        
        If fi > 0 And fo > 0 Then
            astr = ":alpha='if(lt(t," & st & "),0,if(lt(t," & st + fi & "),(t-" & st & ")/" & fi & ",if(lt(t," & et - fo & ")," & al & ",if(lt(t," & et & "),(" & fo & "-(t-" & et - fo & "))/" & fo & ",0))))'"
        End If
        
        If fi > 0 And fo = 0 Then
            astr = ":alpha='if(lt(t," & st & "),0,if(lt(t," & st + fi & "),(t-" & st & ")/" & fi & ",if(lt(t," & et & ")," & al & ")))'"
        End If
        
        If fi = 0 And fo > 0 Then
            astr = ":alpha='if(lt(t," & et - fo & ")," & al & ",if(lt(t," & et & "),(" & fo & "-(t-" & et - fo & "))/" & fo & ",0))'"
        End If
        
        If fi = 0 And fo = 0 Then
            astr = ":alpha=" & al
        End If
    
        cmd = cmd & ":x=" & oTextDraw.XLoc & ":y=" & oTextDraw.YLoc & ":font=" & oTextDraw.Font & ":fontsize=" & oTextDraw.FontSize & ":fontcolor=" & oTextDraw.FontColor '& ":style=" & fs
        If oTextDraw.Box = 1 Then
            cmd = cmd & ":box=" & oTextDraw.Box
            cmd = cmd & ":boxcolor=" & oTextDraw.BoxColor
            cmd = cmd & ":boxborderw=" & oTextDraw.BoxBorderWeight
        End If
        If oTextDraw.LineSpacing > 0 Then cmd = cmd & ":line_spacing=" & oTextDraw.LineSpacing
        If oTextDraw.BorderWeight > 0 Then cmd = cmd & ":borderw=" & oTextDraw.BorderWeight
        If LCase(oTextDraw.BorderColor) <> "black" Then cmd = cmd & ":bordercolor=" & oTextDraw.BorderColor
        cmd = cmd & astr
        If oTextDraw.ShadowX <> 0 Or oTextDraw.ShadowY <> 0 Then
            cmd = cmd & ":shadowx=" & oTextDraw.ShadowX & ":shadowy=" & oTextDraw.ShadowY & ":shadowcolor=" & oTextDraw.ShadowColor
        End If
        If oTextDraw.TabSize <> 4 Then cmd = cmd & ":tabsize=" & oTextDraw.TabSize
        
    Next i
    
    cmd = cmd & """ -c:a copy" & encodeStr & " -y " & outstream

    SendCommand cmd
End Sub

Public Function Probe() As ffProbe
    Dim p As New ffProbe
    p.DefaultIOPath = mDefaultIOPath
    Set Probe = p
End Function

Public Function GetProgramVersion() As String
    Dim jc As New JSonConverter
    cmd = "-hide_banner -loglevel error -v 0 -of default=nw=1:nk=1 -show_program_version -print_format json"
    resp = Me.Probe.RunCommand(cmd)
    GetProgramVersion = jc.ParseJSON(resp)("program_version")("version")
End Function

Private Function GetAbsolutePath(ByVal strPath As String, Optional ByVal refPath As String = "") As String
    Dim fso As New FileSystemObject
    saveppath = CurDir()
    If refPath = "" Then refPath = ThisWorkbook_PathOnDisk
    ChDrive refPath
    ChDir refPath
    GetAbsolutePath = fso.GetAbsolutePathName(VBA.Trim(strPath))
    ChDrive saveppath
    ChDir saveppath
End Function

Public Function ThisWorkbook_PathOnDisk() As String
    ' The reason for this function is that when the workbook is opened on a disk synched with OneDrive or SharePoint,
    '  (ThisWorkbook.FullName and) ThisWorkbook.Path returns the correspondent cloud URLs instead than the original path on disk. For example:
    ' "https://d.docs.live.net/e06a[etc...]/MyDocumentFolder/MyFolder"
    ' or "https://mycompany.sharepoint.com/personal/MyName_Company_com/MyDocumentFolder/mycompany/Apps/BlaBla"
    ' causing problems if that path is used with other functions, like ChDrive.
    '
    ' This function must be used as a replacement to "ThisWorkbook.Path" that always returns the original/real path on disk. For example:
    ' "C:\Users\myUserName\OneDrive\Documenti\MyFolder"

    Dim strPath As String
    strPath = ThisWorkbook.path

    If VBA.Left$(strPath, 8) = "https://" Then
        'Original script taken from https://stackoverflow.com/a/72736800/11738627  (credits to GWD and his sources)
        Const HKEY_CURRENT_USER = &H80000001
        Dim objReg As WbemScripting.SWbemObjectEx 'changed to early binding by GCUser99
        Dim regPath As String
        Dim subKeys() As Variant
        Dim subKey As Variant
        Dim strValue As String
        Dim strMountpoint As String
        Dim strSecPart As String
        
        Static pathSep As String
        If pathSep = "" Then pathSep = Application.PathSeparator
    
        Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

        regPath = "Software\SyncEngines\Providers\OneDrive\"
        objReg.EnumKey HKEY_CURRENT_USER, regPath, subKeys
    
        For Each subKey In subKeys
            objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "UrlNamespace", strValue
            If InStr(strPath, strValue) > 0 Then
                objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "MountPoint", strMountpoint
                strSecPart = Replace(Mid(strPath, Len(strValue)), "/", pathSep)
                ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
    
                Do Until Dir(ThisWorkbook_PathOnDisk, vbDirectory) <> "" Or InStr(2, strSecPart, pathSep) = 0
                    strSecPart = Mid(strSecPart, InStr(2, strSecPart, pathSep))
                    ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
                Loop
                Exit Function
            End If
        Next subKey
    End If
        
    ThisWorkbook_PathOnDisk = strPath
    
End Function



Private Function TimeStringToSecs(ByVal ts As String) As Double
    tsparts = Split(ts, ":")
    TimeStringToSecs = 0
    j = 0
    For i = UBound(tsparts) To 0 Step -1
        TimeStringToSecs = TimeStringToSecs + tsparts(i) * (60 ^ j)
        j = j + 1
    Next i
End Function

Private Sub Class_Initialize()
    commandWindowStyle = vbNormalFocus 'vbHide 'hides the command window
    ffmpegExecPath = ".\ffmpeg.exe"
    tmpDirPath = Environ("TEMP") 'same as fso.GetSpecialFolder(2)
    mDefaultIOPath = GetAbsolutePath(".\")
End Sub

Private Function GetInfo(ByVal inputFilepath As String) As ffStreamInfo
    Dim jc As New JSonConverter
    Dim d As Dictionary
    Dim streamInfo As ffStreamInfo
    Dim streams As Collection

    commandStr = "-hide_banner -print_format json -show_entries stream=codec_name,codec_type,duration,pix_fmt,width,height"
    
    resp = Me.Probe.RunCommand(commandStr, inputFilepath)
    
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

Private Function GetXfadeTransString(ByVal val As ffXFadeTrans) As String
    Select Case val
    Case ffXFadeTrans.ffFade
        GetXfadeTransString = "fade"
    Case ffXFadeTrans.ffWipeleft
        GetXfadeTransString = "wipeleft"
    Case ffXFadeTrans.ffWiperight
        GetXfadeTransString = "wiperight"
    Case ffXFadeTrans.ffWipeup
        GetXfadeTransString = "wipeup"
    Case ffXFadeTrans.ffWipedown
        GetXfadeTransString = "wipedown"
    Case ffXFadeTrans.ffSlideleft
        GetXfadeTransString = "slideleft"
    Case ffXFadeTrans.ffSlideright
        GetXfadeTransString = "slideright"
    Case ffXFadeTrans.ffSlideup
        GetXfadeTransString = "slideup"
    Case ffXFadeTrans.ffSlidedown
        GetXfadeTransString = "slidedown"
    Case ffXFadeTrans.ffCirclecrop
        GetXfadeTransString = "circlecrop"
    Case ffXFadeTrans.ffRectcrop
        GetXfadeTransString = "rectcrop"
    Case ffXFadeTrans.ffDistance
        GetXfadeTransString = "distance"
    Case ffXFadeTrans.ffFadeblack
        GetXfadeTransString = "fadeblack"
    Case ffXFadeTrans.ffFadewhite
        GetXfadeTransString = "fadewhite"
    Case ffXFadeTrans.ffRadial
        GetXfadeTransString = "radial"
    Case ffXFadeTrans.ffSmoothleft
        GetXfadeTransString = "smoothleft"
    Case ffXFadeTrans.ffSmoothright
        GetXfadeTransString = "smoothright"
    Case ffXFadeTrans.ffSmoothup
        GetXfadeTransString = "smoothup"
    Case ffXFadeTrans.ffSmoothdown
        GetXfadeTransString = "smoothdown"
    Case ffXFadeTrans.ffCircleopen
        GetXfadeTransString = "circleopen"
    Case ffXFadeTrans.ffCircleclose
        GetXfadeTransString = "circleclose"
    Case ffXFadeTrans.ffVertopen
        GetXfadeTransString = "vertopen"
    Case ffXFadeTrans.ffVertclose
        GetXfadeTransString = "vertclose"
    Case ffXFadeTrans.ffHorzopen
        GetXfadeTransString = "horzopen"
    Case ffXFadeTrans.ffHorzclose
        GetXfadeTransString = "horzclose"
    Case ffXFadeTrans.ffDissolve
        GetXfadeTransString = "dissolve"
    Case ffXFadeTrans.ffPixelize
        GetXfadeTransString = "pixelize"
    Case ffXFadeTrans.ffDiagtl
        GetXfadeTransString = "diagtl"
    Case ffXFadeTrans.ffDiagtr
        GetXfadeTransString = "diagtr"
    Case ffXFadeTrans.ffDiagbl
        GetXfadeTransString = "diagbl"
    Case ffXFadeTrans.ffDiagbr
        GetXfadeTransString = "diagbr"
    Case ffXFadeTrans.ffHlslice
        GetXfadeTransString = "hlslice"
    Case ffXFadeTrans.ffHrslice
        GetXfadeTransString = "hrslice"
    Case ffXFadeTrans.ffVuslice
        GetXfadeTransString = "vuslice"
    Case ffXFadeTrans.ffVdslice
        GetXfadeTransString = "vdslice"
    Case ffXFadeTrans.ffHblur
        GetXfadeTransString = "hblur"
    Case ffXFadeTrans.ffFadegrays
        GetXfadeTransString = "fadegrays"
    Case ffXFadeTrans.ffWipetl
        GetXfadeTransString = "wipetl"
    Case ffXFadeTrans.ffWipetr
        GetXfadeTransString = "wipetr"
    Case ffXFadeTrans.ffWipebl
        GetXfadeTransString = "wipebl"
    Case ffXFadeTrans.ffwipebr
        GetXfadeTransString = "wipebr"
    Case ffXFadeTrans.ffSqueezeh
        GetXfadeTransString = "squeezeh"
    Case ffXFadeTrans.ffSqueezev
        GetXfadeTransString = "squeezev"
    Case ffXFadeTrans.ffZoomin
        GetXfadeTransString = "zoomin"
    End Select
End Function

Private Function GetFilterString(cf As ffCurvesFilter) As String
    haveone = False
    cmd = "curves="
    
    If cf.Preset <> ffNone Then
        cmd = cmd & "preset=" & GetCurvesPresetString(cf.Preset)
        haveone = True
    End If
    If cf.All <> "" Then
        If haveone Then cmd = cmd & ":"
        cmd = cmd & "all='" & cf.All & "'"
        haveone = True
    End If
    If cf.Master <> "" Then
        If haveone Then cmd = cmd & ":"
        cmd = cmd & "master='" & cf.Master & "'"
        haveone = True
    End If
    If cf.Psfile <> "" Then
        If haveone Then cmd = cmd & ":"
        'cmd = cmd & "psfile=""" & cf.Psfile & """"
        cmd = cmd & "psfile='" & EscapePath(GetAbsolutePath(cf.Psfile, mDefaultIOPath)) & "'"
        haveone = True
    End If
    If cf.Plot <> "" Then
        If haveone Then cmd = cmd & ":"
        'cmd = cmd & "plot=""" & cf.Plot & """"
        cmd = cmd & "plot='" & EscapePath(GetAbsolutePath(cf.Plot, mDefaultIOPath)) & "'"
        haveone = True
    End If
    If cf.Red <> "" Then
        If haveone Then cmd = cmd & ":"
        cmd = cmd & "red='" & cf.Red & "'"
        haveone = True
    End If
    If cf.Blue <> "" Then
        If haveone Then cmd = cmd & ":"
        cmd = cmd & "blue='" & cf.Blue & "'"
        haveone = True
    End If
    If cf.Green <> "" Then
        If haveone Then cmd = cmd & ":"
        cmd = cmd & "green='" & cf.Green & "'"
        haveone = True
    End If

    GetFilterString = cmd
End Function

Private Function GetCurvesPresetString(ByVal curvesPreset As ffCurvesPresets) As String
    Select Case curvesPreset
    Case ffCurvesPresets.ffColorNegative
        GetCurvesPresetString = "color_negative"
    Case ffCurvesPresets.ffCrossProcess
        GetCurvesPresetString = "cross_process"
    Case ffCurvesPresets.ffDarker
        GetCurvesPresetString = "darker"
    Case ffCurvesPresets.ffIncreaseContrast
        GetCurvesPresetString = "increase_contrast"
    Case ffCurvesPresets.ffLighter
        GetCurvesPresetString = "lighter"
    Case ffCurvesPresets.ffLinearContrast
        GetCurvesPresetString = "linear_contrast"
    Case ffCurvesPresets.ffMediumContrast
        GetCurvesPresetString = "medium_contrast"
    Case ffCurvesPresets.ffNegative
        GetCurvesPresetString = "negative"
    Case ffCurvesPresets.ffNone
        GetCurvesPresetString = "none"
    Case ffCurvesPresets.ffStrongContrast
        GetCurvesPresetString = "strong_contrast"
    Case ffCurvesPresets.ffVintage
        GetCurvesPresetString = "vintage"
    End Select
End Function

Private Function EscapePath(ByVal path As String) As String
    'some filters require : and \ chars in filepath to be escaped with \ - example: C\:\\User\\Documents\\filter.acv
    EscapePath = Replace(Replace(path, "\", "\\"), ":", "\:")
End Function

Private Function EscapeSpecialChars(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\\\")
    txt = Replace(txt, "'", "\\'")
    txt = Replace(txt, ":", "\:")
    txt = Replace(txt, "%", "\\%")
    txt = Replace(txt, "{", "\\{")
    EscapeSpecialChars = Replace(txt, "}", "\\}")
End Function

Private Function is2d(a As Variant) As Boolean
    Dim l As Long
    On Error Resume Next
    l = LBound(a, 2)
    is2d = Err = 0
End Function

Private Function IsImage(ByVal filepath As String) As Boolean
    Dim fso As New FileSystemObject
    ext = fso.GetExtensionName(filepath)
    extList = ".jpg.jpeg.jpe.jif.jfif.jfi"
    extList = extList & ".png"
    extList = extList & ".tiff.tif"
    extList = extList & ".webp"
    extList = extList & ".raw.arw.cr2.nrw.k25"
    extList = extList & ".bmp.dib"
    extList = extList & ".heif.heic"
    extList = extList & ".ind.indd.indt"
    extList = extList & ".jp2.j2k.jpf.jpx.jpm.mj2"
    extList = extList & ".gif"
    extList = extList & "."
    If InStr(1, extList, "." & ext & ".", vbTextCompare) Then
        IsImage = True
    Else
        IsImage = False
    End If
End Function
