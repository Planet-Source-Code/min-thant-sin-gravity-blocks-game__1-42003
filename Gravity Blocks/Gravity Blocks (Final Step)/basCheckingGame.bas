Attribute VB_Name = "basCheckingGame"
Option Explicit

Private colCheckBlocks As New Collection

'The logic here is we find a group of same blocks based on every block
'and add them to the collection. If the number of blocks found is
'equal or greater than MinBlocksToClick (minimum blocks to click
'in order to remove them) , we know there are more blocks to remove
'and the game is NOT over.

Public Function GameOver() As Boolean
      Dim I As Integer, J As Integer
      
      GameOver = True
      'Check all the blocks left
      For I = 1 To colBlocksLeft.Count
            Set colCheckBlocks = Nothing
            Set colCheckBlocks = New Collection
            
            Call CheckBlock(colBlocksLeft(I))
            
            'Reset the found flag to False
            For J = 1 To colCheckBlocks.Count
                  Blocks(colCheckBlocks(J)).HasBeenFound = False
            Next J
            
            'There are still blocks to remove
            If colCheckBlocks.Count >= MinBlocksToClick Then
                  GameOver = False
                  Exit Function
            End If
            
      Next I
End Function

'This sub finds a group of same blocks and add them to the
'colCheckBlocks collection.
Sub CheckBlock(ByVal Index As Integer)
      Dim I As Integer, J As Integer
      Dim DiffX As Integer, DiffY As Integer
      
      If colCheckBlocks.Count >= MinBlocksToClick Then Exit Sub
      
      For I = 1 To colBlocksLeft.Count
            J = colBlocksLeft(I)
            'Same number and this block hasn't been found yet.
            If (Blocks(J).Number = Blocks(Index).Number) And (Blocks(J).HasBeenFound = False) Then
                  'Same column location.
                  If Blocks(J).XCoord = Blocks(Index).XCoord Then
                        'Are they adjacent (touching)?
                        DiffY = Abs(Blocks(J).YCoord - Blocks(Index).YCoord)
                        If DiffY = 1 Then
                              'Indicate this block has been found.
                              Blocks(J).HasBeenFound = True
                              'Add it to the collection
                              colCheckBlocks.Add Blocks(J).ID, Blocks(J).Key
                              'Find others based on this one
                              Call CheckBlock(Blocks(J).ID)
                        End If
                  Else
                        'The same as above...
                        If Blocks(J).YCoord = Blocks(Index).YCoord Then
                              DiffX = Abs(Blocks(J).XCoord - Blocks(Index).XCoord)
                              If DiffX = 1 Then
                                    Blocks(J).HasBeenFound = True
                                    colCheckBlocks.Add Blocks(J).ID, Blocks(J).Key
                                    Call CheckBlock(Blocks(J).ID)
                              End If
                        End If
                  End If
            End If
      Next I
End Sub
