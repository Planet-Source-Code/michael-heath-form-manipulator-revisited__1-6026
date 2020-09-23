VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3990
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frDisclaimer 
      Caption         =   "Disclaimer..."
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   30
      TabIndex        =   4
      Top             =   2730
      Width           =   5205
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblDisclaimer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disclaimer......"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   3750
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frDescription 
      Caption         =   "App Version and Description"
      ForeColor       =   &H8000000D&
      Height          =   2235
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   5205
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   1185
         Left            =   120
         TabIndex        =   3
         Top             =   870
         Width           =   4935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AppName and Version"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   2070
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5235
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
frmAd.Show
Unload Me
End Sub

Private Sub Form_Load()
lblCaption.width = frmAbout.width
lblCaption.Top = 0
lblCaption.Left = 0
frmOnTop Me, True
frmPosition Me, 1
End Sub
