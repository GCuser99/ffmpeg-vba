VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffCurvesFilter"
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
Public Enum ffCurvesPresets
    ffNone
    ffColorNegative
    ffCrossProcess
    ffDarker
    ffIncreaseContrast
    ffLighter
    ffLinearContrast
    ffMediumContrast
    ffNegative
    ffStrongContrast
    ffVintage
End Enum

Public Enum ffCurves
    ffAll
    ffMaster
    ffRed
    ffGreen
    ffBlue
    ffPsfile
    ffPlot
End Enum

Private mPreset As ffCurvesPresets
Private mRed As String
Private mGreen As String
Private mBlue As String
Private mAll As String
Private mMaster As String
Private mPsfile As String
Private mPlot As String

Public Property Let Red(ByVal val As String)
    mRed = val
End Property

Public Property Get Red() As String
    Red = mRed
End Property

Public Property Let Green(ByVal val As String)
    mGreen = val
End Property

Public Property Get Green() As String
    Green = mGreen
End Property

Public Property Let Blue(ByVal val As String)
    mBlue = val
End Property

Public Property Get Blue() As String
    Blue = mBlue
End Property

Public Property Let All(ByVal val As String)
    mAll = val
End Property

Public Property Get All() As String
    All = mAll
End Property

Public Property Let Master(ByVal val As String)
    mMaster = val
End Property

Public Property Get Master() As String
    Master = mMaster
End Property

Public Property Let Psfile(ByVal val As String)
    'takes a PhotoShop acv file path
    mPsfile = val
End Property

Public Property Get Psfile() As String
    Psfile = mPsfile
End Property

Public Property Let Plot(ByVal val As String)
    'Gnuplot script of the curves
    mPlot = val
End Property

Public Property Get Plot() As String
    Plot = mPlot
End Property

Public Property Let Preset(val As ffCurvesPresets)
    mPreset = val
End Property

Public Property Get Preset() As ffCurvesPresets
    Preset = mPreset
End Property

Private Sub Class_Initialize()
    mPreset = ffNone
    mRed = ""
    mGreen = ""
    mBlue = ""
    mAll = ""
    mMaster = ""
    mPsfile = ""
    mPlot = ""
End Sub

Public Sub Reset()
    Class_Initialize
End Sub
