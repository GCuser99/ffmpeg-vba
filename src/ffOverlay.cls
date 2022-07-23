VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffOverlay"
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
Private mInputPath As String
Private mResize As Single
Private mXLoc As String
Private mYLoc As String
Private mStartTime As Single
Private mEndTime As Single
Private mFadeInDuration As Single
Private mFadeOutDuration As Single
Private mCropX As Integer
Private mCropY As Integer
Private mCropWidth As Integer
Private mCropHeight As Integer

Public Property Let InputPath(ByVal val As String)
    mInputPath = val
End Property
Public Property Get InputPath() As String
    InputPath = mInputPath
End Property

Public Property Let Resize(ByVal val As Single)
    mResize = val
End Property
Public Property Get Resize() As Single
    Resize = mResize
End Property

Public Property Let XLoc(ByVal val As String)
    'x location in pixels for left edge of text
    'to center use syntax "(w-text_w)/2"
    val = LCase(val)
    If InStr(val, "center") Then val = Replace(val, "center", "(W-overlay_w)/2")
    If InStr(val, "right") Then val = Replace(val, "right", "(W-overlay_w)")
    If InStr(val, "left") Then val = Replace(val, "left", "0")
    mXLoc = val
End Property
Public Property Get XLoc() As String
    XLoc = mXLoc
End Property

Public Property Let YLoc(ByVal val As String)
    'y location in pixels for upper edge of text
    'to center use syntax "(h-text_h)/2"
    val = LCase(val)
    If InStr(val, "center") Then val = Replace(val, "center", "(H-overlay_h)/2")
    If InStr(val, "bottom") Then val = Replace(val, "bottom", "(H-overlay_h)")
    If InStr(val, "top") Then val = Replace(val, "top", "0")
    mYLoc = val
End Property
Public Property Get YLoc() As String
    YLoc = mYLoc
End Property

Public Property Let startTime(ByVal val As Single)
    mStartTime = val
End Property
Public Property Get startTime() As Single
    startTime = mStartTime
End Property

Public Property Let endTime(ByVal val As Single)
    mEndTime = val
End Property
Public Property Get endTime() As Single
    endTime = mEndTime
End Property

Public Property Let FadeInDuration(ByVal val As Single)
    mFadeInDuration = val
End Property
Public Property Get FadeInDuration() As Single
    FadeInDuration = mFadeInDuration
End Property

Public Property Let FadeOutDuration(ByVal val As Single)
    mFadeOutDuration = val
End Property
Public Property Get FadeOutDuration() As Single
    FadeOutDuration = mFadeOutDuration
End Property

Public Sub SetCrop(Optional ByVal x As Integer, Optional ByVal y As Integer, Optional ByVal width As Integer, Optional ByVal height As Integer)
    mCropX = x
    mCropY = y
    mCropWidth = width
    mCropHeight = height
End Sub

Private Sub Class_Initialize()
    mResize = 1#
End Sub

