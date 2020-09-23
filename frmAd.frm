VERSION 5.00
Begin VB.Form frmAd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advertisement - Make Some Money"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6525
      Begin VB.TextBox txtAd 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000001&
         Height          =   1455
         Left            =   270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   5985
      End
      Begin VB.Image imgAd 
         Height          =   600
         Left            =   270
         Picture         =   "frmAd.frx":0000
         Top             =   390
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmAd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open App.Path & "\frmad.mhx" For Input As 1
    txtAd.Text = Input$(LOF(1), 1)
    Close 1
    LoadNextToMain Me, frmMain, vTopCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmShrink Me
End Sub

Private Sub imgAd_Click()
nResult = Shell("Start.exe http://www.bepaid.com/user.rhtml?REFID=10088516", vbHide)
End Sub
