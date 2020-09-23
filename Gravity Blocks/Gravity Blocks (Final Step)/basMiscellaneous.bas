Attribute VB_Name = "basMiscellaneous"
Option Explicit

Public Sub PlaySound()
      If Dir(App.Path & "\drop.wav") = "" Then Exit Sub
      If FileLen(App.Path & "\drop.wav") > 4000 Then Exit Sub
      
      sndPlaySound App.Path & "\drop.wav", SND_ASYNC
End Sub

'As the name suggests...
Public Sub Delay(ByVal sec As Single)
      Dim Marker As Single
            
      Marker = Timer
      Do Until Timer > Marker + sec
            DoEvents
      Loop
End Sub

Public Sub CleanUpCollections()
      Set colSameBlocks = Nothing
      Set colBlocksLeft = Nothing
      Set colBlocksRemoved = Nothing
End Sub
