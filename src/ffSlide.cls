VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffSlide"
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
Public Enum ffXFadeTrans
    ffFade
    ffWipeleft
    ffWiperight
    ffWipeup
    ffWipedown
    ffSlideleft
    ffSlideright
    ffSlideup
    ffSlidedown
    ffCirclecrop
    ffRectcrop
    ffDistance
    ffFadeblack
    ffFadewhite
    ffRadial
    ffSmoothleft
    ffSmoothright
    ffSmoothup
    ffSmoothdown
    ffCircleopen
    ffCircleclose
    ffVertopen
    ffVertclose
    ffHorzopen
    ffHorzclose
    ffDissolve
    ffPixelize
    ffDiagtl
    ffDiagtr
    ffDiagbl
    ffDiagbr
    ffHlslice
    ffHrslice
    ffVuslice
    ffVdslice
    ffHblur
    ffFadegrays
    ffWipetl
    ffWipetr
    ffWipebl
    ffwipebr
    ffSqueezeh
    ffSqueezev
    ffZoomin
End Enum

Private mInputPath As String
Private mDuration As Single
Private mTransitionDuration As Single
Private mTransitionType As ffXFadeTrans

Public Property Let TransitionType(ByVal val As ffXFadeTrans)
    mTransitionType = val
End Property

Public Property Get TransitionType() As ffXFadeTrans
    TransitionType = mTransitionType
End Property

Public Property Let InputPath(ByVal val As String)
    mInputPath = val
End Property

Public Property Get InputPath() As String
    InputPath = mInputPath
End Property

Public Property Let Duration(ByVal val As Single)
    mDuration = val
End Property
Public Property Get Duration() As Single
    Duration = mDuration
End Property

Public Property Let TransitionDuration(ByVal val As Single)
    mTransitionDuration = val
End Property
Public Property Get TransitionDuration() As Single
    TransitionDuration = mTransitionDuration
End Property

Private Sub Class_Initialize()
    mTransitionType = ffFade
End Sub


