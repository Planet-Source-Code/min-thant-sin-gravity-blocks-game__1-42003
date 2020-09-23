VERSION 5.00
Begin VB.Form frmPuzzle 
   BackColor       =   &H00808000&
   Caption         =   "Gravity Blocks   (Any bugs? Feel free to e-mail me at <minsin999@hotmail.com>)"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   5475
      Picture         =   "GravityBlocks.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picBlocks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   75
      Picture         =   "GravityBlocks.frx":37182
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   4650
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   750
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
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game...      "
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPlaySound 
         Caption         =   "Play &Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu sepExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "How to play..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sepAboutMe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////////////////
'///// Gravity Blocks by Min Thant Sin /////
'/////          December 29, 2002           /////
'///////////////////////////////////////

'Any bugs, comments, suggestions?
'Feel free to e-mail me at:
'< minsin999@hotmail.com >

Public boolProcessing As Boolean    'Indicate moving and dropping of blocks being taken placed

Private Sub Form_Load()
      boolProcessing = False
      boolPlaySound = True
      
      'Init block width & height variables
      BlockWidth = frmPuzzle.picBlank.ScaleWidth
      BlockHeight = frmPuzzle.picBlank.ScaleHeight
      
      Call InitPuzzle(4, 2, 15, 10)
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Call CleanUpCollections
End Sub

Private Sub mnuAbout_Click()
      'It is not necessary.
      If boolProcessing Then Exit Sub
      MsgBox "Created by Min Thant Sin on Dec 29, 2002", vbInformation
End Sub

Private Sub mnuExit_Click()
      Call CleanUpCollections
      End
End Sub

Private Sub mnuHowToPlay_Click()
      'It is not necessary.
      If boolProcessing Then Exit Sub
      frmHowToPlay.Show vbModal
End Sub

Private Sub mnuNewGame_Click()
      'It is not necessary.
      If boolProcessing Then Exit Sub
      frmNewGame.Show vbModal
End Sub

Private Sub mnuPlaySound_Click()
      mnuPlaySound.Checked = Not mnuPlaySound.Checked
      boolPlaySound = mnuPlaySound.Checked
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If boolGameOver Then Exit Sub
      
      'All the moving and dropping of the blocks being taken place.
      'Doesn't allow any action while it is being taken place.
      If boolProcessing Then Exit Sub
      
      'Indicate this is happening...
      boolProcessing = True
      
      '/////////////////////////////////////////////////////////////
      'Retrieving the block index under the mouse and checking if the
      'index is out of range. And if it is not outta range, check if the
      'block exists.
      '/////////////////////////////////////////////////////////////
      
      Dim Index As Integer    'Block index
      
      'Get the block index under the mouse
      Index = GetBlockUnderMouse(X, Y)
      
      'There are two reasons that we don't take action:
      '(1) Index is out of range, it is because the mouse X or Y is negative
      '(2) The block being clicked (mouse up) doesn't exist.
      
      '-1 means out of range
      If Index = -1 Then
            'Indicate no action is being taken place
            boolProcessing = False
            Exit Sub    'No more processing
      Else
            'The user clicked on blank block.
            If Blocks(Index).Exists = False Then
                  'Indicate no action is being taken place
                  boolProcessing = False
                  Exit Sub    'No more processing
            End If
      End If
      
      '////////////////////////////////////////////////////////
      'Finding the same blocks like the clicked block and checking
      'if they can be removed.
      '////////////////////////////////////////////////////////
      
      'Reset same blocks collection
      Set colSameBlocks = Nothing
      Set colSameBlocks = New Collection
      
      'Find same blocks using recursive function
      Call FindSameBlocks(Blocks(Index).ID)
      
      Dim I As Integer
      'Under minimum blocks to click?
      If colSameBlocks.Count < MinBlocksToClick Then
            For I = 1 To colSameBlocks.Count
                  Index = colSameBlocks(I)
                  'Reset found flag to False for next time search
                  Blocks(Index).HasBeenFound = False
            Next I
            'Indicate no action is being taken place
            boolProcessing = False
            Exit Sub    'No more processing
      End If
            
      '//////////////////////////////////////////////////////////
      'Removing the blank blocks from the collection, adding them to
      'another collection, and displaying blank image in their locations.
      '//////////////////////////////////////////////////////////
      
      'Remove a group of same blocks from the collection and
      'set their flag to indicate they no longer exist. Also add
      'them to the collection named < colBlocksRemoved >
      For I = 1 To colSameBlocks.Count
            Index = colSameBlocks(I) 'Retrieve index
            Blocks(Index).Exists = False  'Set flag
            'Remove it from the collection
            colBlocksLeft.Remove Blocks(Index).Key
            'Add it to the collection
            colBlocksRemoved.Add Blocks(Index).ID, Blocks(Index).Key
      Next I
            
      Dim J As Byte
      
      'Display animated blank image.
      For J = 0 To 4
            For I = 1 To colSameBlocks.Count
                  'Get index and...
                  Index = colSameBlocks(I)
                  
                  'Blit that location with animation graphic
                  Call BitBlt(picBoard.hDC, _
                        Blocks(Index).Left, Blocks(Index).Top, _
                        BlockWidth, BlockHeight, picFade.hDC, _
                        J * BlockWidth, Blocks(Index).Number * BlockHeight, _
                        vbSrcCopy)
            Next I
                  Delay 0.05
                  picBoard.Refresh
      Next J
      
      
      '///////////////////////////////////////////////////////////////////////
      'Moving the blocks to bottom, in other words, dropping the blocks
      'that are above the blank blocks.
      '///////////////////////////////////////////////////////////////////////
      
      'To indicate that there is/are blocks to drop
      Dim boolBlocksToDrop As Boolean
      'The block(s) above the removed block(s)
      Dim AboveBlock As Integer
      
      'The logic here is that we find all the blocks that are "adjacent"
      '(touching) AND "above" the removed (blank) blocks .
      'Then we drop them down simultaneously.
      Do
            boolBlocksToDrop = False
            'Check all the blank blocks
            For I = 1 To colSameBlocks.Count
                  'Retrieve index from collection
                  Index = colSameBlocks(I)
                  'Get the above block's index
                  AboveBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord - 1)
                  '-1 means out of range. It occurs when the blank block
                  'is on top. In other words, its YCoord is 0 and you see
                  'that when we subtract 1 from its YCoord. There is no
                  'coordinate as -1. Don't confuse this -1 with the return
                  'value of the above function, which in this case is -1.
                  If AboveBlock <> -1 Then
                        'We drop the block only it exists, it's obvious.
                        If Blocks(AboveBlock).Exists Then
                              'Indicate that there is/are blocks to drop
                              boolBlocksToDrop = True
                              'Now drop them
                              Call Drop(AboveBlock)
                        End If
                  End If
            Next I
                  
            DoEvents
            'Delay dropping, kinda animation
            Call Delay(0.1)
            picBoard.Refresh
            
      'More blocks to drop? Loop again.
      Loop While boolBlocksToDrop
      
      
      
      '///////////////////////////////////////////////////////////
      'Moving the blocks that are behind the gap to the left
      '///////////////////////////////////////////////////////////
      
      Dim RightBlock As Integer     'Blank block's right block
      Dim BaseBlock As Integer      'Bottom-most block (if this block is blank, this indicates a gap)
      Dim BlankBlock As Integer     'Blank block we're checking from left to right direction
      Dim OldXCoord As Integer     'The base block's X Coord
      Dim Row As Integer, Col As Integer  'Row and column
      Dim colBlocksToMove As New Collection     'Blocks to shift to the left
      
      'The following statements seem to be rather lengthy and
      'cumbersome and of course, complicated. But this is only
      'my algorithm, and you could always create your own.
      'In fact, if you go step by step and examine the code
      'carefully, you'll be amazed that it's pretty easy.
      
      'We check from PuzzleWidth - 2 because PuzzleWidth - 1 is
      'the last column and there can be no blocks after this.
      For I = PuzzleWidth - 2 To 0 Step -1
            'Reset collection
            Set colBlocksToMove = Nothing
            Set colBlocksToMove = New Collection
            
            'Base (bottom-most) block
            BaseBlock = Board(I, PuzzleHeight - 1)
            'False means it is a gap.
            If Blocks(BaseBlock).Exists = False Then
                  'Mark this blank block's X coordinate
                  OldXCoord = Blocks(BaseBlock).XCoord
                  
                  'Check from bottom to top and...
                  For Row = PuzzleHeight - 1 To 0 Step -1
                        'Left to right direction
                        For Col = 0 To PuzzleWidth - 1
                              BlankBlock = GetBlockFromCoord(Col, Row)
                              'If this block is blank and its XCoord is equal or
                              'greater than that of base block...
                              If Blocks(BlankBlock).Exists = False Then
                                    If Blocks(BlankBlock).XCoord >= OldXCoord Then
                                          RightBlock = GetBlockFromCoord(Blocks(BlankBlock).XCoord + 1, Blocks(BlankBlock).YCoord)
                                          '-1 means this block is at the last column and there
                                          'is no block to exchange.
                                          If RightBlock <> -1 Then
                                                'If the right block exists, exchange it with
                                                'the blank block.
                                                If Blocks(RightBlock).Exists Then
                                                      Call MoveRightUntilNoMoreBlocks(Blocks(BlankBlock).ID, Blocks(RightBlock).ID)
                                                End If
                                          End If
                                    End If
                              End If
                        Next Col
                  Next Row
            End If
            picBoard.Refresh
      Next I
      
      picBoard.Refresh
      'Indicate the action has been taken.
      boolProcessing = False
      
      If GameOver Then
            boolGameOver = True
            frmGameOver.Show vbModal
      End If
End Sub
