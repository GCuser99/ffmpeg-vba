VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffSlideShow"
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
Private colSlides As Collection

Public Sub MakeSlides(ByVal numSlides As Integer)
    Dim s As ffSlide
    For i = 1 To numSlides
        Set s = New ffSlide
        colSlides.Add s
    Next i
End Sub

Public Function Item(ByVal slideItem As Integer) As ffSlide
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colSlides(slideItem)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemID = -4
    'uncomment line above in text editor to enable For Each iteration in for loop
    Set NewEnum = colSlides.[_NewEnum]
End Function

Public Function Count() As Integer
    Count = colSlides.Count
End Function

Public Property Let TransitionType(ByVal val As ffXFadeTrans)
    For i = 1 To Count
        Item(i).TransitionType = val
    Next i
End Property

Public Property Let Duration(ByVal val As Single)
    For i = 1 To Count
        Item(i).Duration = val
    Next i
End Property

Public Property Let TransitionDuration(ByVal val As Single)
    For i = 1 To Count
        Item(i).TransitionDuration = val
    Next i
End Property

Private Sub Class_Initialize()
    Set colSlides = New Collection
End Sub

