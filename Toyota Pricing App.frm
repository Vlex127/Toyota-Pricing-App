VERSION 5.00
Begin VB.Form CSC_207_GROUP_3 
   Caption         =   "Toyota Pricing App"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   Icon            =   "Toyota Pricing App.frx":0000
   LinkTopic       =   "Toyota Pricing App"
   ScaleHeight     =   8355
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Info"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   2760
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtDuty 
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtYear 
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtEngine 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtMake 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Text            =   "Toyota"
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Cost"
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CheckBox chkRoof 
      Caption         =   "Open Roof"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CheckBox chkAC 
      Caption         =   "Air Conditioning"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblResult 
      Caption         =   "€ 0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   3000
      TabIndex        =   11
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Import Duty %"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Year Manufactured"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Engine Size"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Make 
      Caption         =   "Make"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "CSC_207_GROUP_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
    Dim ListPage1 As String
    Dim ListPage2 As String
    
    ' --- PAGE 1 (First 18 Names) ---
    ListPage1 = "CSC 207 - Group 3 (Page 1 of 2)" & vbCrLf & _
                "--------------------------------" & vbCrLf & _
                "240591081 – Eniola Uthman Oluwadarasimi" & vbCrLf & _
                "240591082 – EZE C.O" & vbCrLf & _
                "240591083 – Ezebube Chizitelum Odirachukwunma" & vbCrLf & _
                "240591084 – Ezeh Johncharles" & vbCrLf & _
                "240591085 – Olawale Anuoluwapo Ezekiel" & vbCrLf & _
                "240591087 – Fajemisin Alexander Olumuyiwa" & vbCrLf & _
                "240591088 – Fakoya Khalid" & vbCrLf & _
                "240591089 – Familusi Ayomikun" & vbCrLf & _
                "240591090 – Fasasi Mayowa" & vbCrLf & _
                "240591091 – Fasoyiro Samuel Oreoluwa" & vbCrLf & _
                "240591092 – Giwa Adesope Abdulmalik" & vbCrLf & _
                "240591093 – Giwa Aliameen Opeyemi" & vbCrLf & _
                "240591094 – Marvellous Godwin Osemudiamen" & vbCrLf & _
                "240591095 – Hammed Abdulrahman Adesina" & vbCrLf & _
                "240591096 – Hanidu Olawale Murtadha" & vbCrLf & _
                "240591097 – Hassan Mubarak Atanda" & vbCrLf & _
                "240591098 – Segun Ganiu Hassan" & vbCrLf & _
                "240591101 – Idowu Fawaz Olawunmi" & vbCrLf & vbCrLf & _
                ">>> Click OK for Next Page >>>"

    ' --- PAGE 2 (Remaining Names) ---
    ListPage2 = "CSC 207 - Group 3 (Page 2 of 2)" & vbCrLf & _
                "--------------------------------" & vbCrLf & _
                "240591102 – Idowu Matthew" & vbCrLf & _
                "240591103 – Ikechukwu Queen" & vbCrLf & _
                "240591104 – Inedu Esther Ogaicha" & vbCrLf & _
                "240591105 – Iniotuh Greatest" & vbCrLf & _
                "240591106 – Iwakun Oluwasegun Omotojumi" & vbCrLf & _
                "240591107 – Iwuno Vincent ChukwuEbuka" & vbCrLf & _
                "240591108 – Iyaniwura Olabamiji George" & vbCrLf & _
                "240591109 – Jawando Fuad Olamide" & vbCrLf & _
                "240591110 – Jeremiah David Preye" & vbCrLf & _
                "240591111 – Joseph Elizabeth Nifemi" & vbCrLf & _
                "240591112 – Kalejaiye Halimah Temilade" & vbCrLf & _
                "240591113 – Kasali Damilola Emmanuel" & vbCrLf & _
                "240591115 – Kehinde Oyindamola Ayomide" & vbCrLf & _
                "240591116 – Kolawole Abubakar Olaoluwa" & vbCrLf & _
                "240591117 – Kwegan Sean Oluwatomilade" & vbCrLf & _
                "240591118 – Lamina Rihanat Opemipo" & vbCrLf & _
                "240591119 – Lawal Sahal Adeshayo" & vbCrLf & _
                "240591120 – Ibrahim Oluwafemi Lawal"

    ' Display Page 1
    MsgBox ListPage1, vbInformation, "Group 3 List (1/2)"
    
    ' Display Page 2 immediately after they close Page 1
    MsgBox ListPage2, vbInformation, "Group 3 List (2/2)"
    
End Sub

Private Sub cmdCalculate_Click()
    ' 1. Declare variables
    Dim BasePrice As Currency
    Dim EngineSize As Double
    Dim YearMade As Integer
    Dim DutyPercent As Double
    Dim FacilitiesCost As Currency
    Dim TotalCost As Currency
    
    ' --- SAFETY CHECKS (VALIDATION) ---
    ' If the user typed "ABC" or left it blank, stop here to prevent a crash.
    If IsNumeric(txtEngine.Text) = False Then
        MsgBox "Please enter a valid number for Engine Size!", vbExclamation
        txtEngine.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtYear.Text) = False Then
        MsgBox "Please enter a valid Year!", vbExclamation
        txtYear.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtDuty.Text) = False Then
        MsgBox "Please enter a valid number for Duty %!", vbExclamation
        txtDuty.SetFocus
        Exit Sub
    End If
    ' ----------------------------------

    ' 2. Set Base Price
    BasePrice = 5000000
    
    ' 3. Get inputs (Safe to do now because we checked them above)
    EngineSize = Val(txtEngine.Text)
    YearMade = Val(txtYear.Text)
    DutyPercent = Val(txtDuty.Text)
    
    ' 4. Engine Logic
    If EngineSize > 2 Then
        BasePrice = BasePrice + 500000
    End If
    
    ' 5. Year Logic (Updated to ignore "0" / empty box)
    If YearMade > 0 And YearMade < 2015 Then
        BasePrice = BasePrice - 200000
    End If
    
    ' 6. Facilities Logic
    FacilitiesCost = 0
    If chkAC.Value = 1 Then
        FacilitiesCost = FacilitiesCost + 50000
    End If
    If chkRoof.Value = 1 Then
        FacilitiesCost = FacilitiesCost + 150000
    End If
    
    ' 7. Calculate Total
    TotalCost = BasePrice + FacilitiesCost
    TotalCost = TotalCost + (TotalCost * (DutyPercent / 100))
    
    ' 8. Display Result with Euros
    lblResult.Caption = "€ " & Format(TotalCost, "#,##0.00")

End Sub
Private Sub cmdClear_Click()
    ' 1. Clear all text boxes
    txtMake.Text = "Toyota"  ' Reset make to default
    txtEngine.Text = ""
    txtYear.Text = ""
    txtDuty.Text = ""
    
    ' 2. Uncheck the boxes (0 means Unchecked)
    chkAC.Value = 0
    chkRoof.Value = 0
    
    ' 3. Reset the result label
    lblResult.Caption = "€ 0.00"
    
    ' 4. Put the cursor back in the first box for the user
    txtEngine.SetFocus
End Sub

