Attribute VB_Name = "basDisplayingBlocks"
Option Explicit

'Blit the block's image (used when the block moves, drops)
Public Sub BlitImage(ByVal index As Integer)
      Call BitBlt(frmPuzzle.picBoard.hDC, _
            Blocks(index).Left, Blocks(index).Top, _
            BlockWidth, BlockHeight, _
            frmPuzzle.picBlocks.hDC, _
            Blocks(index).Number * BlockWidth, 0, _
            vbSrcCopy)
End Sub

'Blit blank image (used when the block moves, drops or is removed)
Public Sub BlitBlank(ByVal index As Integer)
      Call BitBlt(frmPuzzle.picBoard.hDC, _
            Blocks(index).Left, Blocks(index).Top, _
            BlockWidth, BlockHeight, _
            frmPuzzle.picBlank.hDC, _
            0, 0, _
            vbSrcCopy)
End Sub
