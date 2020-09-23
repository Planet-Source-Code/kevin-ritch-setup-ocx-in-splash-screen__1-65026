VERSION 5.00
Begin VB.Form SplashForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "This is the SPLASH SCREEN"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SplashForm.frx":0000
   ScaleHeight     =   3495
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer SetupTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written for PSC Users by Kevin Ritch (Kevin@V8Software.com)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   6450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modules written with my RESOURCE BUILDER at PSC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   5550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Automatically installs the following components"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   4875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Widget Software for Geeks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   5745
   End
End
Attribute VB_Name = "SplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'=======================================================
'WARNING - AFTER RUNNING, CHECK TO SEE IF YOUR
'PROJECT COMPONENTS NEED REASSIGNING TO WINDOWS\SYSTEM32
'=======================================================
'The Modules for the resources were all wriiten with
'Resource Builder I already uploaded to PSC
'============================================================
'The REGISTRATION MODULE I LIFTED - FORGET WHERE - Maybe PSC!
'============================================================
 SetupTimer.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
 TheMainForm.Show
End Sub
Private Sub SetupTimer_Timer()
 SetupTimer.Enabled = False
 Call SetupComponents
End Sub
Private Sub Timer1_Timer()
 Unload Me
End Sub
Sub SetupComponents()
'================================================================
'Be clever, of course - stick components in an array or something
'Maybe even use the PAINT command for a Percentage Bar
'Also don't forget to remind the user to rather SAVE the App
'onto their Desktop for ease of execution in the future.
'The method used here simply installs the components...
'And can be INSTANTLY EXECUTED from a ZIP FILE over the Internet
'Feel at ease to FREELY COPY and use this source code - ENJOY!!!!
'================================================================
 On Error Resume Next
 MkDir "c:\Program Files\"
 MkDir "c:\Program Files\Widget_Software_Demo_2006\"
 On Error GoTo 0
'=====================
'COMMON DIALOG CONTROL
'=====================
 List1.AddItem "COMDLG32.ocx"
 List1.Refresh
 If Dir$("c:\Program Files\Widget_Software_Demo_2006\COMDLG32.ocx") = "" Then
  Call BuildComDLG32_ocx
 '============================================================
 'SUGGEST REM OFF OCX REGISTRATION DURING SOFTWARE DEVELOPMENT
 '============================================================
  Result = VBRegSvr32("c:\Program Files\Widget_Software_Demo_2006\COMDLG32.ocx", True)
 End If
'===========
'TAB CONTROL
'===========
 List1.AddItem "TabCTL32.ocx"
 List1.Refresh
 If Dir$("c:\Program Files\Widget_Software_Demo_2006\tabctl32.ocx") = "" Then
  Call BuildTabCTL32_OCX
 '============================================================
 'SUGGEST REM OFF OCX REGISTRATION DURING SOFTWARE DEVELOPMENT
 '============================================================
  Result = VBRegSvr32("c:\Program Files\Widget_Software_Demo_2006\TabCTL32.ocx", True)
 End If
'=========================
'MICROSOFT COMMON CONTROLS
'=========================
 List1.AddItem "MSComCTL.ocx"
 List1.Refresh
 If Dir$("c:\Program Files\Widget_Software_Demo_2006\MSComCTL.ocx") = "" Then
  Call BuildMSComCTL_OCX
 '============================================================
 'SUGGEST REM OFF OCX REGISTRATION DURING SOFTWARE DEVELOPMENT
 '============================================================
  Result = VBRegSvr32("c:\Program Files\Widget_Software_Demo_2006\MSComCTL.ocx", True)
 End If
'=========================================================================
'The OCX Registration, if used, changes the NeededComponents Value to True
'=========================================================================
 If NeededComponents = False Then
  Timer1.Enabled = True ' 8 second splash!
 Else
  Unload Me ' It probably took 8 seconds or so!
 End If
End Sub
