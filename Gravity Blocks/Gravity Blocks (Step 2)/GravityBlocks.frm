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

Private Sub Form_Load()
      'Init block width & height variables
      BlockWidth = frmPuzzle.picBlank.ScaleWidth
      BlockHeight = frmPuzzle.picBlank.ScaleHeight
      
      Call InitPuzzle(4, 2, 15, 10)
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim Index As Integer    'Block index
      
      'Get the block index under the mouse
      Index = GetBlockUnderMouse(X, Y)
      If Index = -1 Then Exit Sub
      If Blocks(Index).Exists = False Then Exit Sub
      
      'Reset same blocks collection
      Set colSameBlocks = Nothing
      Set colSameBlocks = New Collection
      
      'Find same blocks using recursive function
      Call FindSameBlocks(Blocks(Index).ID)
      
      Dim I As Integer
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
      
      picBoard.Refresh
End Sub

'As the name suggests...
Public Sub Delay(ByVal sec As Single)
      Dim Marker As Single
            
      Marker = Timer
      Do Until Timer > Marker + sec
            DoEvents
      Loop
End Sub
