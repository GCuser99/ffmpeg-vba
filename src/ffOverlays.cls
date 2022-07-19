VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ffOverlays"
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
Private colOverlays As Collection

Public Sub MakeOverlays(ByVal numOverlays As Integer)
    Dim ol As ffOverlay
    For i = 1 To numOverlays
        Set ol = New ffOverlay
        colOverlays.Add ol
    Next i
End Sub

Public Function Item(ByVal index As Integer) As ffOverlay
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colOverlays(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemID = -4
    'uncomment line above in text editor to enable For Each iteration in for loop
    Set NewEnum = colOverlays.[_NewEnum]
End Function

Public Function Count() As Integer
    Count = colOverlays.Count
End Function

Public Property Let Resize(ByVal val As Single)
    For i = 1 To Count
        Item(i).Resize = val
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

Public Sub SetCrop(Optional ByVal x As Integer, Optional ByVal y As Integer, Optional ByVal width As Integer, Optional ByVal height As Integer)
    For i = 1 To Count
        Item(i).SetCrop x, y, width, height
    Next i
End Sub

Private Sub Class_Initialize()
    Set colOverlays = New Collection
End Sub


