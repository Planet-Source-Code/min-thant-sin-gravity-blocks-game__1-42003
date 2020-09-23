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
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public boolProcessing As Boolean    'Indicate moving and dropping of blocks being taken placed

Private Sub Form_Load()
      boolProcessing = False
      'Init block width & height variables
      BlockWidth = frmPuzzle.picBlank.ScaleWidth
      BlockHeight = frmPuzzle.picBlank.ScaleHeight
      
      Call InitPuzzle(4, 2, 15, 10)
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If boolProcessing Then Exit Sub
      
      'Indicate this is happening...
      boolProcessing = True
      
      Dim Index As Integer    'Block index
      
      'Get the block index under the mouse
      Index = GetBlockUnderMouse(X, Y)
      
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
            
      For I = 1 To colSameBlocks.Count
            Index = colSameBlocks(I) 'Retrieve index
            Blocks(Index).Exists = False  'Set flag
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
      
      'To indicate that there is/are blocks to drop
      Dim boolBlocksToDrop As Boolean
      'The block(s) above the removed block(s)
      Dim AboveBlock As Integer
      
      Do
            boolBlocksToDrop = False
            'Check all the blank blocks
            For I = 1 To colSameBlocks.Count
                  'Retrieve index from collection
                  Index = colSameBlocks(I)
                  'Get the above block's index
                  AboveBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord - 1)
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
      
      picBoard.Refresh
      'Indicate the action has been taken.
      boolProcessing = False
End Sub

'As the name suggests...
Public Sub Delay(ByVal sec As Single)
      Dim Marker As Single
            
      Marker = Timer
      Do Until Timer > Marker + sec
            DoEvents
      Loop
End Sub
