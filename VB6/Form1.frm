VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucAniGif ucAniGif1 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
        ucAniGif1.SizeToFit = True
        ucAniGif1.File = App.Path & "\bananatime.gif"
        ucAniGif1.Play
End Sub
