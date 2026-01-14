VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2640
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Toyota Pricing App"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   600
      Picture         =   "frmSplash.frx":1084A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    ' 1. Stop the timer so it doesn't keep looping
    Timer1.Enabled = False
    
    ' 2. Close this splash screen
    Unload Me
    
    ' 3. Open the Main Calculator (Form1)
    CSC_207_GROUP_3.Show
End Sub
