Attribute VB_Name = "basMovingBlocks"
Option Explicit

Public Sub Drop(ByVal Index As Integer)
      Dim BottomBlock As Integer
      Dim temp As Integer
      
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      'Out of range. We check this for safety.
      If BottomBlock = -1 Then Exit Sub
      
      Call BlitBlank(Index)
      
      temp = Blocks(Index).ID
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord) = Blocks(BottomBlock).ID
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord + 1) = temp
      
      Blocks(Index).YCoord = Blocks(Index).YCoord + 1
      Blocks(BottomBlock).YCoord = Blocks(BottomBlock).YCoord - 1
      
      Blocks(Index).Top = Blocks(Index).YCoord * BlockHeight
      Blocks(BottomBlock).Top = Blocks(BottomBlock).YCoord * BlockHeight
      
      Call BlitImage(Index)
End Sub
