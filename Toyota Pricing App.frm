VERSION 5.00
Begin VB.Form CSC_207_GROUP_3 
   BackColor       =   &H00F8F9FA&
   Caption         =   "Toyota Pricing App"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H006C757D&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DC3545&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtDuty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtYear 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtEngine 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtMake 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   7
      Text            =   "Toyota"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0028A745&
      Caption         =   "Calculate Cost"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CheckBox chkRoof 
      BackColor       =   &H00F8F9FA&
      Caption         =   "Open Roof (+N50,000)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox chkAC 
      BackColor       =   &H00F8F9FA&
      Caption         =   "Air Conditioning (+N75,000)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N 0.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0028A745&
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Import Duty (%)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00495057&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Manufactured"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00495057&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Engine Size (cc)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00495057&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Make 
      BackStyle       =   0  'Transparent
      Caption         =   "Make"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00495057&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblTotalCost 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00495057&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00DEE2E6&
      Height          =   6855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   7095
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Pricing Calculator"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB0000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   4695
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
    Dim ListPage3 As String
      
    ' --- PAGE 1 (Usage Instructions + First 10 Names) ---
    ' FIX: Split into multiple lines AND replaced bullets with * or + symbols
    ListPage1 = "TOYOTA PRICING APP - USAGE INSTRUCTIONS" & vbCrLf & _
                "==========================================" & vbCrLf
    ListPage1 = ListPage1 & "HOW TO USE:" & vbCrLf
    ListPage1 = ListPage1 & "1. Enter Engine Size (cc) - e.g., 2000" & vbCrLf
    ListPage1 = ListPage1 & "2. Enter Year Manufactured - e.g., 2020" & vbCrLf
    ListPage1 = ListPage1 & "3. Enter Import Duty % - e.g., 35" & vbCrLf
    ListPage1 = ListPage1 & "4. Select optional features:" & vbCrLf
    ListPage1 = ListPage1 & "   + Air Conditioning (+N75,000)" & vbCrLf
    ListPage1 = ListPage1 & "   + Open Roof (+N50,000)" & vbCrLf
    ListPage1 = ListPage1 & "5. Click 'Calculate Cost' for total price" & vbCrLf
    ListPage1 = ListPage1 & "6. Click 'Clear' to reset all fields" & vbCrLf & vbCrLf
     
    ListPage1 = ListPage1 & "PRICING LOGIC:" & vbCrLf
    ListPage1 = ListPage1 & "* Small engines (<=2000cc): N1.5M base" & vbCrLf
    ListPage1 = ListPage1 & "* Medium engines (<=3000cc): N2.5M base" & vbCrLf
    ListPage1 = ListPage1 & "* Large engines (>3000cc): N4.0M base" & vbCrLf
    ListPage1 = ListPage1 & "* Newer cars (>=2020): +N500,000" & vbCrLf
    ListPage1 = ListPage1 & "* Older cars (<2015): -N300,000" & vbCrLf & vbCrLf
     
    ListPage1 = ListPage1 & "CSC 207 - Group 3 Members (Page 1 of 3)" & vbCrLf
    ListPage1 = ListPage1 & "----------------------------------------" & vbCrLf
    ListPage1 = ListPage1 & "240591081 - Eniola Uthman Oluwadarasimi" & vbCrLf
    ListPage1 = ListPage1 & "240591082 - EZE C.O" & vbCrLf
    ListPage1 = ListPage1 & "240591083 - Ezebube Chizitelum Odirachukwunma" & vbCrLf
    ListPage1 = ListPage1 & "240591084 - Ezeh Johncharles" & vbCrLf
    ListPage1 = ListPage1 & "240591085 - Olawale Anuoluwapo Ezekiel" & vbCrLf
    ListPage1 = ListPage1 & "240591087 - Fajemisin Alexander Olumuyiwa" & vbCrLf
    ListPage1 = ListPage1 & "240591088 - Fakoya Khalid" & vbCrLf
    ListPage1 = ListPage1 & "240591089 - Familusi Ayomikun" & vbCrLf
    ListPage1 = ListPage1 & "240591090 - Fasasi Mayowa" & vbCrLf
    ListPage1 = ListPage1 & "240591091 - Fasoyiro Samuel Oreoluwa" & vbCrLf & vbCrLf
    ListPage1 = ListPage1 & ">>> Click OK for Next Page >>>"

    ' --- PAGE 2 (Remaining Names) ---
    ListPage2 = "CSC 207 - Group 3 Members (Page 2 of 3)" & vbCrLf & _
                "----------------------------------------" & vbCrLf
    ListPage2 = ListPage2 & "240591092 - Giwa Adesope Abdulmalik" & vbCrLf
    ListPage2 = ListPage2 & "240591093 - Giwa Aliameen Opeyemi" & vbCrLf
    ListPage2 = ListPage2 & "240591094 - Marvellous Godwin Osemudiamen" & vbCrLf
    ListPage2 = ListPage2 & "240591095 - Hammed Abdulrahman Adesina" & vbCrLf
    ListPage2 = ListPage2 & "240591096 - Hanidu Olawale Murtadha" & vbCrLf
    ListPage2 = ListPage2 & "240591097 - Hassan Mubarak Atanda" & vbCrLf
    ListPage2 = ListPage2 & "240591098 - Segun Ganiu Hassan" & vbCrLf
    ListPage2 = ListPage2 & "240591101 - Idowu Fawaz Olawunmi" & vbCrLf
    ListPage2 = ListPage2 & "240591102 - Idowu Matthew" & vbCrLf
    ListPage2 = ListPage2 & "240591103 - Ikechukwu Queen" & vbCrLf
    ListPage2 = ListPage2 & "240591104 - Inedu Esther Ogaicha" & vbCrLf
    ListPage2 = ListPage2 & "240591105 - Iniotuh Greatest" & vbCrLf
    ListPage2 = ListPage2 & "240591106 - Iwakun Oluwasegun Omotojumi" & vbCrLf
    ListPage2 = ListPage2 & "240591107 - Iwuno Vincent ChukwuEbuka" & vbCrLf
    ListPage2 = ListPage2 & "240591108 - Iyaniwura Olabamiji George" & vbCrLf & vbCrLf
    ListPage2 = ListPage2 & ">>> Click OK for Final Page >>>"

    ' --- PAGE 3 (Final Names) ---
    ListPage3 = "CSC 207 - Group 3 Members (Page 3 of 3)" & vbCrLf & _
                "----------------------------------------" & vbCrLf
    ListPage3 = ListPage3 & "240591109 - Jawando Fuad Olamide" & vbCrLf
    ListPage3 = ListPage3 & "240591110 - Jeremiah David Preye" & vbCrLf
    ListPage3 = ListPage3 & "240591111 - Joseph Elizabeth Nifemi" & vbCrLf
    ListPage3 = ListPage3 & "240591112 - Kalejaiye Halimah Temilade" & vbCrLf
    ListPage3 = ListPage3 & "240591113 - Kasali Damilola Emmanuel" & vbCrLf
    ListPage3 = ListPage3 & "240591115 - Kehinde Oyindamola Ayomide" & vbCrLf
    ListPage3 = ListPage3 & "240591116 - Kolawole Abubakar Olaoluwa" & vbCrLf
    ListPage3 = ListPage3 & "240591117 - Kwegan Sean Oluwatomilade" & vbCrLf
    ListPage3 = ListPage3 & "240591118 - Lamina Rihanat Opemipo" & vbCrLf
    ListPage3 = ListPage3 & "240591119 - Lawal Sahal Adeshayo" & vbCrLf
    ListPage3 = ListPage3 & "240591120 - Ibrahim Oluwafemi Lawal" & vbCrLf & vbCrLf
    ListPage3 = ListPage3 & "Thank you for using Toyota Pricing App!" & vbCrLf
    ListPage3 = ListPage3 & "Developed by CSC 207 Group 3 2024"

    ' Display Page 1
    MsgBox ListPage1, vbInformation, "How to Use - Page 1/3"
      
    ' Display Page 2
    MsgBox ListPage2, vbInformation, "Group Members - Page 2/3"
      
    ' Display Page 3
    MsgBox ListPage3, vbInformation, "Group Members - Page 3/3"
      
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
    ' --------------------------------
      
    ' 2. Read values from textboxes
    EngineSize = CDbl(txtEngine.Text)
    YearMade = CInt(txtYear.Text)
    DutyPercent = CDbl(txtDuty.Text)
      
    ' 3. Calculate Base Price based on Engine Size
    If EngineSize <= 2000 Then
        BasePrice = 1500000  ' N1.5M for small engines
    ElseIf EngineSize <= 3000 Then
        BasePrice = 2500000  ' N2.5M for medium engines
    Else
        BasePrice = 4000000  ' N4M for large engines
    End If
      
    ' 4. Add Year-based adjustments
    If YearMade >= 2020 Then
        BasePrice = BasePrice + 500000  ' Newer cars cost more
    ElseIf YearMade < 2015 Then
        BasePrice = BasePrice - 300000  ' Older cars cost less
    End If
      
    ' 5. Calculate Facilities Cost (checkboxes)
    FacilitiesCost = 0
    If chkAC.Value = 1 Then
        FacilitiesCost = FacilitiesCost + 75000  ' AC costs N75,000
    End If
      
    If chkRoof.Value = 1 Then
        FacilitiesCost = FacilitiesCost + 50000  ' Open roof costs N50,000
    End If
      
    ' 6. Calculate Import Duty
    Dim ImportDuty As Currency
    ImportDuty = BasePrice * (DutyPercent / 100)
      
    ' 7. Calculate Total Cost
    TotalCost = BasePrice + ImportDuty + FacilitiesCost
      
    ' 8. Display Result with proper formatting
    lblResult.Caption = "N " & Format(TotalCost, "#,##0.00")
      
End Sub

Private Sub cmdClear_Click()
    txtEngine.Text = ""
    txtYear.Text = ""
    txtDuty.Text = ""
    chkAC.Value = 0
    chkRoof.Value = 0
    lblResult.Caption = "N 0.00"
    txtEngine.SetFocus
End Sub