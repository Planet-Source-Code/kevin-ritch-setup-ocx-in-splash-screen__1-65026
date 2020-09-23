VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form TheMainForm 
   Caption         =   "MAIN FORM - Much easier to deploy because the OCX's are EMBEDED !!!"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   3840
      TabIndex        =   5
      Top             =   2760
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "TheMainForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column3"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test CommonDialog1"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Components were all installed and registered in the Splash Screen!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   6840
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Components used:   COMDLG32.OCX  TABCTL32.OCX   MSCOMCTL.OCX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "TheMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 CommonDialog1.Action = 1
 
End Sub

Private Sub Form_Load()
 ProgressBar1.Value = 67
End Sub
