VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H007E9681&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H007E9681&
      Caption         =   "&OK"
      Height          =   315
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Width           =   1005
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   90
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   75
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Beep
frmKillExit Me
End Sub
