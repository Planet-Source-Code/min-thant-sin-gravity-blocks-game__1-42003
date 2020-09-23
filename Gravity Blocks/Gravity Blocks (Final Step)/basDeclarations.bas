Attribute VB_Name = "basDeclarations"
Option Explicit

Public Const SND_ASYNC& = &H1

'Bit block transfer yada yada yada
Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Public Declare Function sndPlaySound& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)

Public Type GRAVITY_BLOCK
       ID As Integer           'Used as collections' Item
       Number As Integer  'Stores a random number. Based on that number, the block's image change.
       XCoord As Integer   'X coordinate of the block
       YCoord As Integer   'Y coordinate of the block
      
      'Actual location of the block
       Left As Long
       Right As Long
       Top As Long
       Bottom As Long
      
       Key As String    'Used as collections' Key
       Exists As Boolean      'Set to False when removed from the board
       HasBeenFound As Boolean      'Used in finding same blocks. Indicate that this block has
      'already been found. To avoid repetition.
End Type

Public PuzzleWidth As Byte    'Blocks horizontally in a row
Public PuzzleHeight As Byte   'Blocks vertically in a column
Public MinBlocksToClick As Byte        'A group of minimum blocks to click to remove them
Public Const MaxBlocksToClick = 4   'Maximum blocks to click

Public TotalBlocks As Integer    'Total blocks in the board
Public NumBlockTypes As Integer 'Number of block types  (there are 6 block types - A, B, C, D, E, F)
Public BlockWidth As Integer  'A block's width & height
Public BlockHeight As Integer

Public boolPlaySound As Boolean
Public boolGameOver As Boolean

Public Board() As Integer           'To keep track of blocks' positions

Public colBlocksLeft As New Collection         'Remaining blocks
Public colBlocksRemoved As New Collection  'Blocks removed from the board
Public colSameBlocks As New Collection       'A group of same blocks to be removed

Public Blocks() As GRAVITY_BLOCK
