VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffTexts"
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
Private colTexts As Collection

Public Sub MakeTexts(ByVal numTexts As Integer) ', Optional ByVal Duration As Integer = 0, Optional ByVal TransitionDuration As Integer = 0, Optional ByVal TransitionType As ffXFadeTrans = ffFade)
    Dim t As ffText
    For i = 1 To numTexts
        Set t = New ffText
        colTexts.Add t
    Next i
End Sub

Public Function Item(ByVal TextDrawItem As Integer) As ffText
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colTexts(TextDrawItem)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemID = -4
    'uncomment line above in text editor to enable For Each iteration in for loop
    Set NewEnum = colTexts.[_NewEnum]
End Function

Public Function Count() As Integer
    Count = colTexts.Count
End Function

Public Property Let startTime(ByVal val As Single)
    For i = 1 To Count
        Item(i).startTime = val
    Next i
End Property

Public Property Let endTime(ByVal val As Single)
    For i = 1 To Count
        Item(i).endTime = val
    Next i
End Property

Public Property Let Text(ByVal val As String)
    For i = 1 To Count
        Item(i).Text = val
    Next i
End Property

Public Property Let TabSize(ByVal val As Integer)
    For i = 1 To Count
        Item(i).TabSize = val
    Next i
End Property

Public Property Let ShadowColor(ByVal val As String)
    For i = 1 To Count
        Item(i).ShadowColor = val
    Next i
End Property

Public Property Let ShadowX(ByVal val As Integer)
    For i = 1 To Count
        Item(i).ShadowX = val
    Next i
End Property

Public Property Let ShadowY(ByVal val As Integer)
    For i = 1 To Count
        Item(i).ShadowY = val
    Next i
End Property

Public Property Let Alpha(ByVal val As String)
    For i = 1 To Count
        Item(i).Alpha = val
    Next i
End Property

Public Property Let LineSpacing(ByVal val As Integer)
    For i = 1 To Count
        Item(i).LineSpacing = val
    Next i
End Property

Public Property Let FontSize(ByVal val As Integer)
    For i = 1 To Count
        Item(i).FontSize = val
    Next i
End Property

Public Property Let Font(ByVal val As String)
    For i = 1 To Count
        Item(i).Font = val
    Next i
End Property

Public Property Let FontColor(ByVal val As String)
    For i = 1 To Count
        Item(i).FontColor = val
    Next i
End Property

Public Property Let XLoc(ByVal val As String)
    For i = 1 To Count
        Item(i).XLoc = val
    Next i
End Property

Public Property Let YLoc(ByVal val As String)
    For i = 1 To Count
        Item(i).YLoc = val
    Next i
End Property

Public Property Let Box(ByVal val As Boolean)
    For i = 1 To Count
        Item(i).Box = val
    Next i
End Property

Public Property Let BoxBorderWeight(ByVal val As Single)
    For i = 1 To Count
        Item(i).BoxBorderWeight = val
    Next i
End Property

Public Property Let BoxColor(ByVal val As String)
    For i = 1 To Count
        Item(i).BoxColor = val
    Next i
End Property

Public Property Let BorderColor(ByVal val As String)
    For i = 1 To Count
        Item(i).BorderColor = val
    Next i
End Property

Public Property Let BorderWeight(ByVal val As Single)
    For i = 1 To Count
        Item(i).BorderWeight = val
    Next i
End Property

Public Property Let FadeInDuration(ByVal val As Single)
    For i = 1 To Count
        Item(i).FadeInDuration = val
    Next i
End Property

Public Property Let FadeOutDuration(ByVal val As Single)
    For i = 1 To Count
        Item(i).FadeOutDuration = val
    Next i
End Property

Private Sub Class_Initialize()
    Set colTexts = New Collection
End Sub


