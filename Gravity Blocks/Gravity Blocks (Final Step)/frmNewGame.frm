VERSION 5.00
Begin VB.Form frmNewGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New game"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   ControlBox      =   0   'False
   Icon            =   "frmNewGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Puzzle Size"
      Height          =   1065
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2565
      Begin VB.ComboBox cboPuzzleHeight 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   525
         Width           =   840
      End
      Begin VB.ComboBox cboPuzzleWidth 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   525
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Left            =   1500
         TabIndex        =   14
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   1425
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okey-doke"
      Default         =   -1  'True
      Height          =   390
      Left            =   75
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number of block types"
      Height          =   690
      Left            =   75
      TabIndex        =   3
      Top             =   1275
      Width           =   2565
      Begin VB.OptionButton optNumBlockTypes 
         Caption         =   "6"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton optNumBlockTypes 
         Caption         =   "5"
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton optNumBlockTypes 
         Caption         =   "4"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Minimum blocks to click"
      Height          =   690
      Left            =   75
      TabIndex        =   7
      Top             =   2175
      Width           =   2565
      Begin VB.OptionButton optMinBlocksToClick 
         Caption         =   "4"
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton optMinBlocksToClick 
         Caption         =   "3"
         Height          =   240
         Index           =   1
         Left            =   975
         TabIndex        =   9
         Top             =   300
         Width           =   465
      End
      Begin VB.OptionButton optMinBlocksToClick 
         Caption         =   "2"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   8
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tmpMinBlocksToClick As Integer
Private tmpNumBlockTypes As Integer

Private Sub cmdCancel_Click()
      Unload Me
End Sub

Private Sub cmdOK_Click()
      Call InitPuzzle(tmpNumBlockTypes, tmpMinBlocksToClick, Val(cboPuzzleWidth.Text), Val(cboPuzzleHeight.Text))
      Unload Me
End Sub

Private Sub Form_Load()
      optNumBlockTypes(NumBlockTypes - 4).Value = True
      optMinBlocksToClick(MinBlocksToClick - 2).Value = True
      
      Dim I As Integer
      
      For I = 5 To 15
            cboPuzzleWidth.AddItem CStr(I)
            If I <= 10 Then
                  cboPuzzleHeight.AddItem CStr(I)
            End If
      Next I
      
      cboPuzzleWidth.ListIndex = PuzzleWidth - (Val(cboPuzzleWidth.List(0)))
      cboPuzzleHeight.ListIndex = PuzzleHeight - (Val(cboPuzzleHeight.List(0)))
End Sub

Private Sub optMinBlocksToClick_Click(Index As Integer)
      tmpMinBlocksToClick = Val(optMinBlocksToClick(Index).Caption)
End Sub

Private Sub optNumBlockTypes_Click(Index As Integer)
      tmpNumBlockTypes = Val(optNumBlockTypes(Index).Caption)
End Sub
