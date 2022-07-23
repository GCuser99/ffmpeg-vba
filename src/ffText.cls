VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffText"
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
Private mText As String
Private mBox As Integer
Private mBoxBorderWeight As Single 'BoxBorderWeight?
Private mBoxColor As String
Private mLineSpacing As Integer
Private mBorderWeight As Integer 'BorderWeight?
Private mBorderColor As String
Private mFontColor As String
Private mFontsize As Integer
Private mFont As String
Private mAlpha As String
Private mXLoc As String
Private mYLoc As String
'Private mFix_bounds As Boolean
'Private mExpansion As String
Private mStartTime As Single
Private mEndTime As Single
Private mFadeInDuration As Single
Private mFadeOutDuration As Single
Private mShadowX As Integer 'ShadowOffsetX
Private mShadowY As Integer 'ShadowOffsetY
Private mShadowColor As String
Private mTabSize As Integer

Public Property Let TabSize(ByVal val As Integer)
    mTabSize = val
End Property

Public Property Get TabSize() As Integer
    TabSize = mTabSize
End Property

Public Property Let ShadowColor(ByVal val As String)
    mShadowColor = formatcolor(val)
End Property

Public Property Get ShadowColor() As String
    ShadowColor = mShadowColor
End Property

Public Property Let ShadowX(ByVal val As Integer)
    mShadowX = val
End Property

Public Property Get ShadowX() As Integer
    ShadowX = mShadowX
End Property

Public Property Let ShadowY(ByVal val As Integer)
    mShadowY = val
End Property

Public Property Get ShadowY() As Integer
    ShadowY = mShadowY
End Property

Private Sub Class_Initialize()
    mBox = 0
    mBoxColor = "white"
    mBoxBorderWeight = 0
    mLineSpacing = 0
    mBorderWeight = 0
    mBorderColor = "black"
    mAlpha = 1
    'mExpansion = "normal" 'or "none"
    mFix_bounds = True '??
    mStartTime = 0
    mEndTime = -1
    Me.XLoc = "center"
    Me.YLoc = "center"
    mFont = "sans"
    mFontColor = "black"
    mFontsize = 16
    mFadeOutDuration = 0
    mFadeInDuration = 0
    mShadowColor = "black"
    mShadowX = 0
    mShadowY = 0
    mTabSize = 4
End Sub

Public Property Let Alpha(ByVal val As Single)
    'between 0.0 and 1.0
    mAlpha = val
End Property

Public Property Get Alpha() As Single
    Alpha = mAlpha
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

Public Property Let LineSpacing(ByVal val As Integer)
    mLineSpacing = val
End Property

Public Property Get LineSpacing() As Integer
    LineSpacing = mLineSpacing
End Property

Public Property Let FontSize(ByVal val As Integer)
    mFontsize = val
End Property

Public Property Get FontSize() As Integer
    FontSize = mFontsize
End Property

Public Property Let Font(ByVal val As String)
    mFont = val
End Property

Public Property Get Font() As String
    Font = mFont
End Property

Public Property Let FontColor(ByVal val As String)
    'examples:
    'RGB(0, 0, 255)
    'RGB(0, 0, 255) & "@0.5"
    '"red"
    '"red@0.5"
    '"#0000ff80"
    '"#0000ff@0x80"
    '"0x0000ff@0x80"
    mFontColor = formatcolor(val)
End Property

Public Property Get FontColor() As String
    FontColor = mFontColor
End Property

Public Property Let XLoc(ByVal val As String)
    'x location in pixels for left edge of text
    'to center use syntax "(W-text_w)/2"
    val = LCase(val)
    If InStr(val, "center") Then val = Replace(val, "center", "(W-text_w)/2")
    If InStr(val, "right") Then val = Replace(val, "right", "(W-text_w)")
    If InStr(val, "left") Then val = Replace(val, "left", "0")
    mXLoc = val
End Property

Public Property Get XLoc() As String
    XLoc = mXLoc
End Property

Public Property Let YLoc(ByVal val As String)
    'y location in pixels for upper edge of text
    'to center use syntax "(H-text_h)/2"
    val = LCase(val)
    If InStr(val, "center") Then val = Replace(val, "center", "(H-text_h)/2")
    If InStr(val, "bottom") Then val = Replace(val, "bottom", "(H-text_h)")
    If InStr(val, "top") Then val = Replace(val, "top", "0")
    mYLoc = val
End Property

Public Property Get YLoc() As String
    YLoc = mYLoc
End Property

Public Property Let Text(ByVal val As String)
    mText = val
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Let Box(ByVal val As Boolean)
    If val Then mBox = 1 Else mBox = 0
End Property

Public Property Get Box() As Boolean
    If mBox = 1 Then Box = True Else Box = False
End Property

Public Property Let BoxBorderWeight(ByVal val As Single)
    mBoxBorderWeight = val
End Property

Public Property Get BoxBorderWeight() As Single
    BoxBorderWeight = mBoxBorderWeight
End Property

Public Property Let BoxColor(ByVal val As String)
    'examples:
    'RGB(0, 0, 255)
    'RGB(0, 0, 255) & "@0.5"
    '"red"
    '"red@0.5"
    '"#0000ff80"
    '"#0000ff@0x80"
    '"0x0000ff@0x80"
    mBoxColor = formatcolor(val)
End Property

Public Property Get BoxColor() As String
    BoxColor = mBoxColor
End Property

Public Property Let BorderColor(ByVal val As String)
    'examples:
    'RGB(0, 0, 255)
    'RGB(0, 0, 255) & "@0.5"
    '"red"
    '"red@0.5"
    '"#0000ff80"
    '"#0000ff@0x80"
    '"0x0000ff@0x80"
    BorderColor = formatcolor(val)
End Property

Public Property Get BorderColor() As String
    BorderColor = mBorderColor
End Property

Public Property Let BorderWeight(ByVal val As Single)
    mBorderWeight = val
End Property

Public Property Get BorderWeight() As Single
    BorderWeight = mBorderWeight
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

Private Function GetHexFromLong(ByVal longColor As Long) As String

    Dim r As String
    Dim g As String
    Dim b As String
    
    r = Format(Application.WorksheetFunction.Dec2Hex(longColor Mod 256), "00")
    g = Format(Application.WorksheetFunction.Dec2Hex((longColor \ 256) Mod 256), "00")
    b = Format(Application.WorksheetFunction.Dec2Hex((longColor \ 65536) Mod 256), "00")
    
    'r = Format(RGBDec2Hex(longColor Mod 256), "00")
    'g = Format(RGBDec2Hex((longColor \ 256) Mod 256), "00")
    'b = Format(RGBDec2Hex((longColor \ 65536) Mod 256), "00")
    
    GetHexFromLong = "#" & r & g & b

End Function

Private Function RGBDec2Hex(ByVal rgbdec As Long)
    'ms office compatible
    rgbdec = RGB(RGBBlue(rgbdec), RGBGreen(rgbdec), RGBRed(rgbdec)) 'to be compatible with MS office
    RGBDec2Hex = "#" & String(6 - Len(Hex(rgbdec)), "0") & Hex(rgbdec)
End Function

Private Function RGBRed(ByVal rgbcol As Long) As Long
    'Return the Red component from an RGB Color
   RGBRed = rgbcol And &HFF
End Function

Private Function RGBGreen(ByVal rgbcol As Long) As Long
    'Return the Green component from an RGB Color
   RGBGreen = ((rgbcol And &H100FF00) / &H100)
End Function

Private Function RGBBlue(ByVal rgbcol As Long) As Long
    'Return the Blue component from an RGB Color
   RGBBlue = (rgbcol And &HFF0000) / &H10000
End Function

Private Function formatcolor(ByVal val As String) As String
    '[0x|#]RRGGBB[AA], followed by optional @alpha can be hex (0x00-0xff) or decimal 0.0-1.0
    'this also takes an RGB value such as RGB(0,0,255) & "@0.5"
    'can also specify a builtin color like blue@0.5
    'these are all equivalent:
    '"blue@0.5"
    'RGB(0, 0, 255) & "@0.5"
    '"#0000ff80"
    '"#0000ff@0x80"
    '"0x0000ff@0x80"
    '"0x0000ff@0.5"
    
    Dim color As String, Alpha As String
    
    'first separate color from alpha if there is an explicit alpha
    ipos = InStr(val, "@")
    If ipos > 0 Then
        color = Left(val, ipos - 1)
        Alpha = Mid(val, ipos + 1, 999)
    Else
        color = val
        Alpha = ""
    End If
    
    If IsNumeric(color) Then 'we have an RGB long value
        color = GetHexFromLong(color)
        If Alpha = "" Then
            val = color
        Else
            val = color & "@" & Alpha
        End If
    End If
    
    formatcolor = val
    
End Function

