VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Toyota Pricing App"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7800
      Top             =   2640
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00666666&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Label lblSubtitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Professional Vehicle Pricing Solution"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
      Width           =   6600
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Toyota Pricing App"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB0000&
      Height          =   735
      Left            =   900
      TabIndex        =   0
      Top             =   2640
      Width           =   7200
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   3600
      Picture         =   "frmSplash.frx":1084A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   5400
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Width           =   7800
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
