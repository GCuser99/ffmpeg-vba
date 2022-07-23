Attribute VB_Name = "examples"
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
' The following subs test most of the ffmpeg-vba functionality herein.
'
' For these test subs to function unmodified, download ffmpeg from:
'
'    https://github.com/BtbN/FFmpeg-Builds/releases
'    or...
'    https://www.gyan.dev/ffmpeg/builds/
'
' Locate the ffmpeg, ffplay, and ffprobe executables in the same directory as this Excel file.
'
' If you wish to store the executables above in a different directory, then modify the executable paths
' in the respective "Class_Initialize" subs of the FFmpeg, FFplay, and FFprobe classes provided here.
'
' Only one input file is needed for these tests:
'
'    http://commondatastorage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4
'
' Download the above into same directory as this excel file.
'
' ffmpeg-vba supports relative file paths. The default i/o path directory is same as where this Excel file resides.
'
' However, that can be changed by using the following command once before file specification:
'    media.DefaultIOPath=[path to your media files]
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
'    - a comma-separated string of time pairs, for example: "0.0 to 5.0, 00:10.0 to 00:15.0, 00:00:20.0 to 00:00:25.0"
'    - a 1D array of time pair strings, for example: Array("0.0 to 5.0", "00:10.0 to 00:15.0", "00:00:20.0 to 00:00:25.0")
'    - or a 2D array of times, for example: Array(Array(0.0,5.0), Array("00:10.0","00:15.0"), Array("00:00:20.0","00:00:25.0"))
'
' ==========================================================================

Sub test_draw_text()
    Dim media As New ffMpeg
    Dim txts As New ffTexts
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'trim the input video to the desired time window
    media.Trim "BigBuckBunny.mp4", "trim.mp4", 27, 47
    
    txts.MakeTexts 10
    
    txts.XLoc = "center"
    txts.YLoc = "bottom - 100" 'place bottom edge of text 100 pixels above bottom
    txts.Font = "arial"
    txts.FontSize = 48
    txts.FontColor = "white"
    txts.FadeInDuration = 0.5
    txts.FadeOutDuration = 0.5
    
    For i = 1 To txts.Count
        txts(i).Text = "This is text number: " & i
        txts(i).startTime = (i - 1) * 2
        txts(i).endTime = (i - 1) * 2 + 2
    Next i
    
    media.DrawText "trim.mp4", "drawtext.mp4", txts
    
    media.DeleteFiles "trim.mp4"
    
    media.Play "drawtext.mp4", , , , 0.5
End Sub

Sub test_slide_show()
    Dim media As New ffMpeg
    Dim slides As New ffSlideShow
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'slides can be videos or images but must have similar encoding
    
    'extract images every 13.5 secs for slide show
    For i = 1 To 44
        media.ExtractFrame "BigBuckBunny.mp4", "slide" & Format(i, "00") & ".jpg", 13.5 * (i - 1) + 1
    Next i
    
    'initialize slides
    slides.MakeSlides 44
    
    'set global properties for slide deck
    slides.Duration = 2.5
    slides.TransitionDuration = 1
    
    'set properties for each slide
    For i = 1 To 44
        slides(i).InputPath = "slide" & Format(i, "00") & ".jpg"
        'set a different xFade transition
        slides(i).TransitionType = i - 1
    Next i
    
    'compile into slide show
    media.MakeSlideShow slides, "slideshow.mp4"
    
    'delete images used
    media.DeleteFiles "slide*.jpg"
    
    'play the show
    media.Play "slideshow.mp4", , , , 0.5
End Sub

Sub test_overlays()
    Dim ovls As New ffOverlays
    Dim media As New ffMpeg
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'overlays can be videos or images or a mix thereof
    
    'extract overlay images/videos from video
    media.Trim "BigBuckBunny.mp4", "overlay01.mp4", 37, 47, True
    media.Trim "BigBuckBunny.mp4", "overlay02.mp4", 67, 77, True
    media.Trim "BigBuckBunny.mp4", "overlay03.mp4", 87, 97, True
    
    'initialize overlays with ffOverlay class
    ovls.MakeOverlays 3
    
    'set some global overlay properties
    ovls.XLoc = "right-10": ovls.YLoc = 10
    ovls.FadeInDuration = 3: ovls.FadeOutDuration = 3
    
    'set individual overlay properties
    ovls(1).InputPath = "overlay01.mp4"
    ovls(1).startTime = 0: ovls(1).endTime = 10
    ovls(1).Resize = 0.3
    
    ovls(2).InputPath = "overlay02.mp4"
    ovls(2).startTime = 7: ovls(2).endTime = 17
    ovls(2).Resize = 0.3
    
    ovls(3).InputPath = "overlay03.mp4"
    ovls(3).startTime = 14: ovls(3).endTime = 24
    ovls(3).Resize = 0.3
    
    'trim the input video
    media.Trim "BigBuckBunny.mp4", "trim.mp4", 27, 52, True
    
    'overlay onto the "base" video
    media.Overlay "trim.mp4", "overlays.mp4", ovls
        
    media.DeleteFiles "trim.mp4", "overlay0*.mp4"
    
    media.Play "overlays.mp4", , , , 0.5
End Sub

Sub test_crop_image()
    Dim media As New ffMpeg
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.ExtractFrame "BigBuckBunny.mp4", "image.jpg", 27
    
    media.Crop "image.jpg", "imagecrop.jpg", 300, 75, 600, 600
    
    media.DeleteFiles "image.jpg"
    
    media.Play "imagecrop.jpg", , , False, 0.5
End Sub

Sub test_run_command()
    Dim media As New ffMpeg
    Dim player As New ffPlay
    Dim prober As New ffProbe
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    'player.DefaultIOPath="[path to your media files]"
    'probe.DefaultIOPath="[path to your media files]"
    
    'RunCommand supports commands that have not been wrapped
    media.RunCommand "-ss 27 -to 33 -c copy -y", "BigBuckBunny.mp4", "trim.mp4"
    
    player.RunCommand "-autoexit -vf ""crop=w=600:h=600:x=300:y=75""", "trim.mp4"
    
    Debug.Print prober.RunCommand("-v error -hide_banner -select_streams v:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1", "trim.mp4")
End Sub

Sub test_curve_filter()
    Dim media As New ffMpeg
    Dim curves As New ffCurvesFilter
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.ExtractFrame "BigBuckBunny.mp4", "image.jpg", 27
    
    curves.All = "0/0 .5/.25 1/1"
    media.CurvesFilter "image.jpg", "image_darker.jpg", curves
    
    curves.All = "0/0 .25/.5 1/1"
    media.CurvesFilter "image_darker.jpg", "image_darker_lighter.jpg", curves
    
    curves.Reset
    curves.Preset = ffColorNegative
    media.CurvesFilter "image.jpg", "image_negative.jpg", curves
    
    curves.Preset = ffStrongContrast
    media.CurvesFilter "image.jpg", "image_strong_contrast.jpg", curves
    
    curves.Reset
    curves.Green = "0/0 .25/.5 1/1"
    media.CurvesFilter "image.jpg", "image_greener.jpg", curves

    'close each of these windows manually
    media.Play "image.jpg", , , False, 0.5
    media.Play "image_darker.jpg", , , False, 0.5
    media.Play "image_darker_lighter.jpg", , , False, 0.5
    media.Play "image_negative.jpg", , , False, 0.5
    media.Play "image_strong_contrast.jpg", , , False, 0.5
    media.Play "image_greener.jpg", , , False, 0.5
End Sub

Sub test_merge_clips()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'break the orig media up into separate pieces - reencode for accuracy
    media.Trim "BigBuckBunny.mp4", "part1.mp4", 27, 32, True
    media.Trim "BigBuckBunny.mp4", "part2.mp4", 32, 37, True
    media.Trim "BigBuckBunny.mp4", "part3.mp4", 37, 42, True
    media.Trim "BigBuckBunny.mp4", "part4.mp4", 42, 47, True
    
    'now join them back up
    media.Join "part1.mp4, part2.mp4, part3.mp4, part4.mp4", "joined.mp4", False
    
    media.DeleteFiles "part1.mp4, part2.mp4, part3.mp4, part4.mp4"
    
    media.Play "joined.mp4", , , , 0.5
End Sub

Sub test_edit()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'edit a media keeping specified time segments
    media.Edit "BigBuckBunny.mp4", "edit.mp4", Array("27 to 32", "37 to 42", "47 to 52", "57 to 62")
    
    media.Play "edit.mp4", , , , 0.5
End Sub

Sub test_reencode()
    Dim media As New ffMpeg
    Dim eSet As New ffEncodeSet
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'try different compression rate factors
    
    eSet.Crf = 25
    media.Trim "BigBuckBunny.mp4", "crf25.mp4", 27, 32, True, , eSet
    
    eSet.Crf = 30
    media.Trim "BigBuckBunny.mp4", "crf30.mp4", 27, 32, True, , eSet
    
    eSet.Crf = 35
    media.Trim "BigBuckBunny.mp4", "crf35.mp4", 27, 32, True, , eSet
    
    'use probe to compare file sizes
    Debug.Print "file size for crf=25 is " & media.Probe.GetFileSize("crf25.mp4") & " mb"
    Debug.Print "file size for crf=30 is " & media.Probe.GetFileSize("crf30.mp4") & " mb"
    Debug.Print "file size for crf=35 is " & media.Probe.GetFileSize("crf35.mp4") & " mb"
    
    'media.DeleteFiles "crf25.mp4", "crf30.mp4", "crf35.mp4"
End Sub

Sub test_meta_rotate()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"

    media.Trim "BigBuckBunny.mp4", "trim.mp4", 27, 47
    
    media.MetaRotate "trim.mp4", "trim_flip_vert_meta.mp4", ffRot180
    
    Debug.Print "meta-rotation= " & media.Probe.GetMetaRotation("trim_flip_vert_meta.mp4")
    
    media.DeleteFiles "trim.mp4"
    
    media.Play "trim_flip_vert_meta.mp4", , , , 0.5
End Sub

Sub test_transpose()
    Dim media As New ffMpeg
    Dim txts As New ffTexts
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.Trim "BigBuckBunny.mp4", "trim.mp4", 27, 47
    
    txts.MakeTexts 1
    txts(1).Text = "This is Left To Right"
    txts(1).Font = "arial"
    txts(1).FontSize = 64
    txts(1).FontColor = "white"
    
    media.DrawText "trim.mp4", "trim_text.mp4", txts
    
    media.Transpose "trim_text.mp4", "trim_flip_lr.mp4", ffFlipLeftRight
    
    Debug.Print "meta-rotation= " & media.Probe.GetMetaRotation("trim_flip_lr.mp4")
    
    media.DeleteFiles "trim.mp4", "trim_text.mp4"
    
    media.Play "trim_flip_lr.mp4", , , , 0.5
End Sub

Sub test_reverse()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.Trim "BigBuckBunny.mp4", "trim.mp4", "1:16", "1:36", True
    
    media.Reverse "trim.mp4", "trim_rev.mp4"
    
    media.DeleteFiles "trim.mp4"
    
    media.Play "trim_rev.mp4", , , , 0.5
End Sub

Sub test_change_speed()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.Trim "BigBuckBunny.mp4", "trim.mp4", "1:16", "1:36", True
    
    media.ChangeSpeed "trim.mp4", "trim_fast.mp4", 4
    
    Debug.Print "input duration: " & media.Probe.GetDuration("trim.mp4")
    Debug.Print "output duration: " & media.Probe.GetDuration("trim_fast.mp4")
    
    media.DeleteFiles "trim.mp4"
    
    media.Play "trim_fast.mp4", , , , 0.5
End Sub

Sub test_dub_over()
    Dim media As New ffMpeg
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    media.Trim "BigBuckBunny.mp4", "trim.mp4", "1:16", "1:36", True
    
    'separate video and audio into two files
    media.RemoveVideo "trim.mp4", "audio_only.m4a"
    
    Debug.Print "has video: " & media.Probe.HasVideo("audio_only.m4a")
    Debug.Print "has audio: " & media.Probe.HasAudio("audio_only.m4a")
    
    media.RemoveAudio "trim.mp4", "video_only.mp4"
    
    Debug.Print "has video: " & media.Probe.HasVideo("video_only.mp4")
    Debug.Print "has audio: " & media.Probe.HasAudio("video_only.mp4")
    
    'now add back audio to video-only file
    media.DubOver "video_only.mp4", "audio_only.m4a", "dub_over.mp4"
    
    Debug.Print "has video: " & media.Probe.HasVideo("dub_over.mp4")
    Debug.Print "has audio: " & media.Probe.HasAudio("dub_over.mp4")
    
    media.DeleteFiles "trim.mp4,audio_only.m4a,video_only.mp4"
    
    media.Play "dub_over.mp4", , , , 0.5
End Sub

Sub test_fade()
    Dim media As New ffMpeg
    Dim oEncSet As New ffEncodeSet
    
    'uncomment and modify commands below if media files are in a different loc than this Excel file
    'media.DefaultIOPath = "[path to your media files]"
    
    'trim the input video to the desired time window
    media.Trim "BigBuckBunny.mp4", "trim.mp4", "1:16", "1:36", True
    
    'reverse the trimmed result
    media.Reverse "trim.mp4", "rev.mp4"
    
    'slow motion of the trimmed result
    media.ChangeSpeed "trim.mp4", "slomo.mp4", 0.5
    
    'join the 3 videos produced above
    media.Join "trim.mp4, rev.mp4, slomo.mp4", "join.mp4"
    
    'now use the Texts class to draw some text overlays
    Dim txts As New ffTexts
    
    txts.MakeTexts 3
    
    'set some global text properties
    txts.XLoc = 10: txts.YLoc = 10
    txts.Font = "arial"
    txts.FontSize = 48
    
    'set individual text properties
    txts(1).Text = "Trimmed Video": txts(1).startTime = 0: txts(1).endTime = 20
    txts(2).Text = "Reversed Video": txts(2).startTime = 20: txts(2).endTime = 40
    txts(3).Text = "Slow Motion Video": txts(3).startTime = 40: txts(3).endTime = 80
    
    'draw the texts onto the composite video
    media.DrawText "join.mp4", "texts.mp4", txts
    
    'specify a constant rate factor for encoding the final result using EncodeSet class
    oEncSet.Crf = 25
    
    'Make 3 second fade from/to black at beginning/end of video
    media.Fade "texts.mp4", "fade.mp4", 3, 3, , oEncSet
    
    'print resulting file size to Intermediate Window
    Debug.Print media.Probe.GetFileSize("fade.mp4") 'in mb
    
    'delete the intermediate file products
    media.DeleteFiles "trim.mp4", "rev.mp4", "slomo.mp4", "join.mp4", "texts.mp4"
    
    'play the result at 50% of the video size
    media.Play "fade.mp4", , , , 0.5
End Sub

Sub test_silence()
    Dim media As New ffMpeg

    media.Trim "BigBuckBunny.mp4", "part1.mp4", "1:16", "1:26", True
    media.Trim "BigBuckBunny.mp4", "part2.mp4", "1:26", "1:36", True
    media.Trim "BigBuckBunny.mp4", "part3.mp4", "1:36", "1:46", True
    
    media.RemoveAudio "part2.mp4", "part2_video_only.mp4"
    
    'if the inputs do not have same number of channels, then the following Join will fail
    'media.Join "part1.mp4, part2_video_only.mp4, part3.mp4", "join_all.mp4"
    
    'so add a silent audio track to part2_video_only.mp4
    
    media.Silence "part2_video_only.mp4", "part2_silence.mp4"
    
    'now join will now succeed!
    media.Join "part1.mp4, part2_silence.mp4, part3.mp4", "join_all.mp4"
    
    media.DeleteFiles "part1.mp4", "part2.mp4", "part3.mp4", "part2_video_only.mp4", "part2_silence.mp4"
    
    media.Play "join_all.mp4", , , , 0.5
End Sub
