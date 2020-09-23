Attribute VB_Name = "basRetrievingData"
Option Explicit

Public Function GetBlockUnderMouse(ByVal X As Single, ByVal Y As Single) As Integer
      Dim XPos As Integer, YPos As Integer
      
      '-1 to indicate mouse position is out of Board range.
      'Mouse position can be negative and there is no Board
      'position such as Board(-8,-5)
      
      GetBlockUnderMouse = -1
      
      'Make sure mouse position is within range
      If X >= 0 And X <= frmPuzzle.picBoard.ScaleWidth Then
            If Y >= 0 And Y <= frmPuzzle.picBoard.ScaleHeight Then
                  'Calculate X & Y coordinates
                  XPos = Int(X / BlockWidth)
                  YPos = Int(Y / BlockHeight)
                  'Return block index that this board position is holding
                  GetBlockUnderMouse = Board(XPos, YPos)
            End If
      End If
End Function

Public Function GetBlockFromCoord(ByVal XCoord As Integer, ByVal YCoord As Integer) As Integer
      '-1 to indicate that it is out of range.
      'We don't have board position such as Board(-1,5).
      GetBlockFromCoord = -1
      
      'Make sure coordinates are within range
      If XCoord >= 0 And XCoord <= PuzzleWidth - 1 Then
            If YCoord >= 0 And YCoord <= PuzzleHeight - 1 Then
                  'Return block index that this board position is holding
                  GetBlockFromCoord = Board(XCoord, YCoord)
            End If
      End If
End Function
