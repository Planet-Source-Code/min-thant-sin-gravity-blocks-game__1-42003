Attribute VB_Name = "basInitializeBlocks"
Option Explicit

Public Sub InitPuzzle(ByVal nBlockTypes As Integer, _
                                    ByVal nBlocksToClick As Integer, _
                                    ByVal nWidth As Integer, _
                                    ByVal nHeight As Integer)
      'Block types (A, B, C, D, E, F)
      NumBlockTypes = nBlockTypes
      
      'Minimum blocks to click to remove them
      MinBlocksToClick = nBlocksToClick
      
      'Init puzzle width & height
      PuzzleWidth = nWidth
      PuzzleHeight = nHeight
            
      ReDim Board(PuzzleWidth - 1, PuzzleHeight - 1) As Integer
            
      'Resize the picBoard
      With frmPuzzle.picBoard
            .Width = .ScaleX(BlockWidth, vbPixels, vbTwips) * PuzzleWidth
            .Height = .ScaleY(BlockHeight, vbPixels, vbTwips) * PuzzleHeight
      End With
      
      With frmPuzzle
      'Center the picBoard
            .picBoard.Left = (Screen.Width - .picBoard.Width) / 2
            .picBoard.Top = (Screen.Height - .picBoard.Height) / 2 - 500
      End With
      'Calculate totalblocks and...
      TotalBlocks = PuzzleWidth * PuzzleHeight
      'Redim blocks
      ReDim Blocks(TotalBlocks - 1) As GRAVITY_BLOCK
      
      Dim X As Integer, Y As Integer, Index As Integer
      
      Call CleanUpCollections
      
      Set colBlocksLeft = New Collection
      Set colBlocksRemoved = New Collection
      Set colSameBlocks = New Collection
      
      'Index is from 0 to TotalBlocks-1
      Index = 0
            
      'Down
      For Y = 0 To PuzzleHeight - 1
            'Across
            For X = 0 To PuzzleWidth - 1
                  Board(X, Y) = Index
                  Randomize Timer
                  
                  With Blocks(Index)
                        .ID = Index
                        .Number = Int(Rnd * NumBlockTypes)
                        .Key = Chr(.Number + 65) & CStr(.ID)
                        .Exists = True
                        .HasBeenFound = False
                        'Coordinates on the board
                        .XCoord = X
                        .YCoord = Y
                        
                        'Actual location on the board
                        .Left = .XCoord * BlockWidth
                        .Top = .YCoord * BlockHeight
                        .Right = .Left + BlockWidth
                        .Bottom = .Top + BlockHeight
                  End With
                  
                  'Display it on the board
                  Call BlitImage(Index)
                  
                  'Store how many blocks there are
                  colBlocksLeft.Add Blocks(Index).ID, Blocks(Index).Key
                  
                  Index = Index + 1
            Next X
      Next Y
            
      boolGameOver = False
      frmPuzzle.picBoard.Refresh
End Sub
