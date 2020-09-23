VERSION 5.00
Begin VB.Form frmPuzzle 
   BackColor       =   &H00808000&
   Caption         =   "Gravity Blocks   (Any bugs? Feel free to e-mail me at <minsin999@hotmail.com>)"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBlocks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   75
      Picture         =   "GravityBlocks.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   75
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   354
      TabIndex        =   0
      Top             =   900
      Width           =   5340
   End
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Bit block transfer yada yada yada
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)

Private Type GRAVITY_BLOCK
      Number As Integer
      XCoord As Integer   'X coordinate of the block
      YCoord As Integer   'Y coordinate of the block
      'Actual location of the block
      Left As Long
      Top As Long
End Type

Private PuzzleWidth As Byte    'Blocks horizontally in a row
Private PuzzleHeight As Byte   'Blocks vertically in a column

Private TotalBlocks As Integer    'Total blocks in the board
Private NumBlockTypes  As Integer
Private BlockWidth As Integer  'A block's width & height
Private BlockHeight As Integer

Private Board() As Integer           'To keep track of blocks' positions
Private Blocks() As GRAVITY_BLOCK

Private Sub Form_Load()
      'Init block width & height variables
      BlockWidth = frmPuzzle.picBlocks.ScaleWidth / 6
      BlockHeight = frmPuzzle.picBlocks.ScaleHeight
      Call InitPuzzle(4, 2, 15, 10)
End Sub

Private Sub InitPuzzle(ByVal nBlockTypes As Integer, _
                                    ByVal nBlocksToClick As Integer, _
                                    ByVal nWidth As Integer, _
                                    ByVal nHeight As Integer)
      'Block types (A, B, C, D, E, F)
      NumBlockTypes = nBlockTypes
      
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
      
      'Index is from 0 to TotalBlocks-1
      Index = 0
            
      'Down
      For Y = 0 To PuzzleHeight - 1
            'Across
            For X = 0 To PuzzleWidth - 1
                  Board(X, Y) = Index
                  Randomize Timer
                  
                  With Blocks(Index)
                        .Number = Int(Rnd * NumBlockTypes)
                        'Coordinates on the board
                        .XCoord = X
                        .YCoord = Y
                        'Actual location on the board
                        .Left = .XCoord * BlockWidth
                        .Top = .YCoord * BlockHeight
                  End With
                  
                  'Display it on the board
                  Call BlitImage(Index)
                  Index = Index + 1
            Next X
      Next Y
            
      frmPuzzle.picBoard.Refresh
End Sub

'Blit the block's image (used when the block moves, drops)
Public Sub BlitImage(ByVal Index As Integer)
      Call BitBlt(frmPuzzle.picBoard.hDC, _
            Blocks(Index).Left, Blocks(Index).Top, _
            BlockWidth, BlockHeight, _
            frmPuzzle.picBlocks.hDC, _
            Blocks(Index).Number * BlockWidth, 0, _
            vbSrcCopy)
End Sub
