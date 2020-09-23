VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Popup Form"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SaveFrmDiminsions Me
End Sub

