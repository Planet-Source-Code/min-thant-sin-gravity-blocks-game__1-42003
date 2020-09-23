Attribute VB_Name = "basDisplayingBlocks"
Option Explicit

'Blit the block's image (used when the block moves, drops)
Public Sub BlitImage(ByVal Index As Integer)
      Call BitBlt(frmPuzzle.picBoard.hDC, _
            Blocks(Index).Left, Blocks(Index).Top, _
            BlockWidth, BlockHeight, _
            frmPuzzle.picBlocks.hDC, _
            Blocks(Index).Number * BlockWidth, 0, _
            vbSrcCopy)
End Sub

Public Sub BlitBlank(ByVal Index As Integer)
      Call BitBlt(frmPuzzle.picBoard.hDC, _
            Blocks(Index).Left, Blocks(Index).Top, _
            BlockWidth, BlockHeight, _
            frmPuzzle.picBlank.hDC, _
            0, 0, _
            vbSrcCopy)
End Sub

