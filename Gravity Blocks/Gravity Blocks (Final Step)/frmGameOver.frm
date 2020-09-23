VERSION 5.00
Begin VB.Form frmGameOver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gravity Blocks"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Okey-dokey!"
      Default         =   -1  'True
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Top             =   2325
      Width           =   2715
   End
   Begin VB.PictureBox picExcellent 
      BackColor       =   &H00000000&
      Height          =   2040
      Left            =   150
      ScaleHeight     =   1980
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   2715
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hey! Excellent!!"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1065
         Left            =   75
         TabIndex        =   13
         Top             =   450
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Game Summary"
      Height          =   2040
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   2715
      Begin VB.Label lblTotalBlocks 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999"
         Height          =   195
         Left            =   1875
         TabIndex        =   11
         Top             =   375
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total blocks:"
         Height          =   195
         Left            =   750
         TabIndex        =   10
         Top             =   375
         Width           =   915
      End
      Begin VB.Label lblNumBlockTypes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999"
         Height          =   195
         Left            =   1875
         TabIndex        =   9
         Top             =   1575
         Width           =   270
      End
      Begin VB.Label lblMinBlocksToClick 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999"
         Height          =   195
         Left            =   1875
         TabIndex        =   8
         Top             =   1275
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Block types:"
         Height          =   195
         Left            =   795
         TabIndex        =   7
         Top             =   1575
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Min blocks to click:"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label lblBlocksRemoved 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999"
         Height          =   195
         Left            =   1875
         TabIndex        =   5
         Top             =   675
         Width           =   270
      End
      Begin VB.Label lblBlocksLeft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999"
         Height          =   195
         Left            =   1875
         TabIndex        =   4
         Top             =   975
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Blocks left:"
         Height          =   195
         Left            =   900
         TabIndex        =   3
         Top             =   975
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Blocks removed:"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   675
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
      picExcellent.Visible = False
      Unload Me
End Sub

Private Sub Form_Load()
      If colBlocksLeft.Count = 0 Then
            picExcellent.Visible = True
            Exit Sub
      End If
      
      lblTotalBlocks.Caption = TotalBlocks
      lblBlocksRemoved.Caption = colBlocksRemoved.Count
      lblBlocksLeft.Caption = colBlocksLeft.Count
      lblMinBlocksToClick.Caption = MinBlocksToClick
      lblNumBlockTypes.Caption = NumBlockTypes
End Sub
