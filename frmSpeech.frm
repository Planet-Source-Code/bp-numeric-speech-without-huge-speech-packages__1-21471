VERSION 5.00
Begin VB.Form frmSpeech 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Speech Example"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Symbol Say"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "1"
      Top             =   0
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Say It!"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Written by:  Blake Pell"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSpeech.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
End
Attribute VB_Name = "frmSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NumericParseTextForSound (Text1.Text)
End Sub

Private Sub Command2_Click()
SymbolParseTextForSound (Text1.Text)
End Sub

Private Sub Form_Load()
ChDir (App.Path)
End Sub

