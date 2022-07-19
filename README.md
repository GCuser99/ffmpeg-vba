## Software: ffmpeg-vba v0.3

This is a wrapper for automating video/image editing with FFmpeg, written in Windows Excel VBA.

Ok, I know, I know... What does this functionality have to do with MS Excel? Not much. However, if like me, you have invested alot in learning the VBA IDE, and you like to build automation templates, then perhaps this media editor might be of interest. Otherwise, this wrapper may be of no use at all... :-)

## Features

- Support for automating command line executables FFmpeg, FFprobe, and FFplay 
- Wrappers for a useful functionality subset, including common tasks like trimming, editing, overlays, filtering, re-encoding, and more
- Convenient video file play-back for instant feedback
- Utility support for relative file paths, VBA RGB color specification, intermediate product file deletion
- Ability to build and run your own commands if wrapper not already provided

## Setup

1) Download/unzip ffmpeg-vba_v0.3.zip to a directory of your choice
2) Download/unzip the [ffpmeg executables](https://ffmpeg.org/download.html) (ffmpeg.exe, ffprobe.exe, ffplay.exe) and place in same directory as the Excel macro file
3) Download the sample video file called "BigBuckBunny.mp4" from [here](http://commondatastorage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4) into same directory
4) Open the excel macro file, go to VBA IDE under the Developer Tab
5) Click on the "examples" standard module to browse and run examples that cover most of the functionality provided

## Example Usage

```vb
Sub example1()
    Dim media As New ffMpeg
    Dim eparms As New ffEncodeSet
    Dim txts As New ffTexts
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
    'trim the input video to the desired time window
    media.Trim "BigBuckBunny.mp4", "trim.mp4", "1:16", "1:36", True
    
    'reverse the trimmed result
    media.Reverse "trim.mp4", "rev.mp4"
    
    'slow motion of the trimmed result
    media.ChangeSpeed "trim.mp4", "slomo.mp4", 0.5
    
    'join the 3 videos produced above
    media.Join "trim.mp4, rev.mp4, slomo.mp4", "join.mp4"
    
    'now use the Texts class to draw some text overlays
    txts.MakeTexts 3
    
    'set some global text properties
    txts.XLoc = 10: txts.YLoc = 10
    txts.Font = "arial"
    txts.FontSize = 48
    
    'set individual text properties
    txts(1).Text = "Trimmed Video": txts(1).StartTime = 0: txts(1).EndTime = 20
    txts(2).Text = "Reversed Video": txts(2).StartTime = 20: txts(2).EndTime = 40
    txts(3).Text = "Slow Motion Video": txts(3).StartTime = 40: txts(3).EndTime = 80
    
    'draw the texts onto the composite video
    media.DrawText "join.mp4", "texts.mp4", txts
    
    'specify a constant rate factor for encoding the final result using EncodeSet class
    eparms.Crf = 25
    
    'make 3 second fade from/to black at beginning/end of video
    media.Fade "texts.mp4", "fade.mp4", 3, 3, , eparms
    
    'print resulting file size to Intermediate Window
    Debug.Print media.Probe.GetFileSize("fade.mp4") 'in mb's
    
    'delete the intermediate file products
    media.DeleteFiles "trim.mp4", "rev.mp4", "slomo.mp4", "join.mp4", "texts.mp4"
    
    'play the result at 50% of the video window size
    media.Play "fade.mp4", , , , 0.5
End Sub
```
```vb
Sub example2()
    Dim media As New ffMpeg
    Dim slides As New ffSlideShow
    
    'uncomment and modify command below if media files are in a different loc than this Excel file
    'media.DefaultIOPath="[path to your media files]"
    
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
    
    'compile into slide show video
    media.MakeSlideShow slides, "slideshow.mp4"
    
    'delete images used with wildcard
    media.DeleteFiles "slide*.jpg"
    
    'play the show
    media.Play "slideshow.mp4", , , , 0.5
End Sub
```

## Collaboration

If you try this and want to report bugs or share ideas for improvement, your contribution is welcome!

## Credits

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA
