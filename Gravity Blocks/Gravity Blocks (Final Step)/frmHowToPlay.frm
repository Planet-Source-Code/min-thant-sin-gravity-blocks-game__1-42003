VERSION 5.00
Begin VB.Form frmHowToPlay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help me please..."
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNotOK 
      Caption         =   "I still don't get it. Help me!"
      Height          =   390
      Left            =   3150
      TabIndex        =   1
      Top             =   3150
      Width           =   2565
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "I got it! I'm now going to play!"
      Default         =   -1  'True
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   3150
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   75
      Picture         =   "frmHowToPlay.frx":0000
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "How many blocks can you remove?"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1440
      Left            =   1575
      TabIndex        =   3
      Top             =   75
      Width           =   4140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmHowToPlay.frx":2CCA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   75
      TabIndex        =   2
      Top             =   1575
      Width           =   5640
   End
End
Attribute VB_Name = "frmHowToPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNotOK_Click()
      Unload Me
      MsgBox "Now, start playing and you'll get it!!", vbInformation
End Sub

Private Sub cmdOK_Click()
      Unload Me
End Sub
