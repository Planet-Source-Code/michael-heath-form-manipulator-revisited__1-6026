VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Main Form"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2760
      Width           =   1000
   End
   Begin VB.Timer tmrNoOffScreen 
      Interval        =   1
      Left            =   5400
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Direction of Opening Form"
      Height          =   1815
      Left            =   1050
      TabIndex        =   0
      Top             =   150
      Width           =   4155
      Begin VB.CommandButton cmdOntop 
         Caption         =   "Advertise"
         Height          =   255
         Index           =   3
         Left            =   2100
         TabIndex        =   19
         Top             =   1410
         Width           =   1000
      End
      Begin VB.CommandButton cmdOntop 
         Caption         =   "CustomMsg"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   1410
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CornerBtmL"
         Height          =   255
         Index           =   11
         Left            =   3060
         TabIndex        =   17
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CornerBtmR"
         Height          =   255
         Index           =   10
         Left            =   2070
         TabIndex        =   16
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CornerTopL"
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   15
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CornerTopR"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   14
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton cmdOntop 
         Caption         =   "NotOnTop"
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   13
         Top             =   1170
         Width           =   1000
      End
      Begin VB.CommandButton cmdOntop 
         Caption         =   "StayOnTop"
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   12
         Top             =   1170
         Width           =   1000
      End
      Begin VB.CommandButton cmdRollUp 
         Caption         =   "Rollup"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1170
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CenterLeft"
         Height          =   255
         Index           =   7
         Left            =   3060
         TabIndex        =   8
         Top             =   510
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CenterRight"
         Height          =   255
         Index           =   6
         Left            =   2070
         TabIndex        =   7
         Top             =   510
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CntrBottom"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Top             =   510
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "CenterTop"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "Bottom"
         Height          =   255
         Index           =   3
         Left            =   3060
         TabIndex        =   4
         Top             =   270
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "Top"
         Height          =   255
         Index           =   2
         Left            =   2070
         TabIndex        =   3
         Top             =   270
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "Left"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   1000
      End
      Begin VB.CommandButton cmdDirection 
         Caption         =   "Right"
         Default         =   -1  'True
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1000
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Try to move this form off your screen to see how the ""NoOffScreen"" sub works."
      Height          =   525
      Left            =   1230
      TabIndex        =   11
      Top             =   2070
      Width           =   3555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsErrorOn 
         Caption         =   "Error Messages O&n"
      End
      Begin VB.Menu mnuOptionsErrorOff 
         Caption         =   "Error Messages O&ff"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuFileRollDown 
         Caption         =   "&RollDown"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpImprovements 
         Caption         =   "&Your Improvements"
      End
      Begin VB.Menu mnuHelpSend 
         Caption         =   "&Send Me Qestions or Comments"
      End
      Begin VB.Menu mnuHelpUpdates 
         Caption         =   "Send &Updates"
      End
      Begin VB.Menu mnuHelpBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDirection_Click(Index As Integer)
    Select Case Index
        Case 0
            'vRight
            LoadNextToMain frmPopup, Me, vRight
            ProgCaption frmPopup, "Loaded on the Right"
        Case 1
            'vLeft
            LoadNextToMain frmPopup, Me, vLeft
            ProgCaption frmPopup, "Loaded on the Left"
        Case 2
            'vTop
            LoadNextToMain frmPopup, Me, vTop
            ProgCaption frmPopup, "Loaded on the Top"
        Case 3
            'vBottom
            LoadNextToMain frmPopup, Me, vBottom
            ProgCaption frmPopup, "Loaded on the Bottom"
        Case 4
            'vTopCenter
            LoadNextToMain frmPopup, Me, vTopCenter
            ProgCaption frmPopup, "Loaded on the Top Centered"
        Case 5
            'vBottomCenter
            LoadNextToMain frmPopup, Me, vBottomCenter
            ProgCaption frmPopup, "Loaded on the  Bottom Centered"
        Case 6
            'vRightCenter
            LoadNextToMain frmPopup, Me, vRightCenter
            ProgCaption frmPopup, "Loaded on the Right Centered"
        Case 7
            'vLeftCenter
            LoadNextToMain frmPopup, Me, vLeftCenter
            ProgCaption frmPopup, "Loaded on the Left Centered"
        Case 8
            ' vCornerTopR
            frmPosition frmPopup, vCornerTopR
                        ProgCaption frmPopup, "Loaded Top Right Corner"

        Case 9
            ' vCornerTopL
            frmPosition frmPopup, vCornerTopL
                        ProgCaption frmPopup, "Loaded Top Left Corner"
            
        Case 10
            ' vCornerBtmR
            frmPosition frmPopup, vCornerBtmR
                        ProgCaption frmPopup, "Loaded Bottom Right Corner"
        
        Case 11
            ' vCornerBtmL
            frmPosition frmPopup, vCornerBtmL
                        ProgCaption frmPopup, "Loaded Bottom Left Corner"
            
    End Select
        
End Sub

Private Sub cmdExit_Click()
'Unload Me
StreakExit Me
End
End Sub


Private Sub cmdOntop_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Put Form on top
            frmOnTop Me, True
        Case 1
            ' Remove Form From being on top
            frmOnTop Me, False
        Case 2
            ' CustomMsg
            CustomMsg "This is a custom message box, and yes I know it sucks." & Chr(10) & "This is your example.  Use it however you want" & Chr(10) & "Now I'm just tryin to take up some space." & Chr(10) & "Taking up space so that you can get the full effect.", "Sucky Message Box", frmIcon.imgIcon
        Case 3
            ' Advertise
            frmAd.Show
    End Select
End Sub

Private Sub cmdRollUp_Click()
mnuView.Visible = True
'DynamicRollUp Me
BooleanRollup Me, True
End Sub

Private Sub Form_Load()
frmCenterMe Me
    frmKillExit Me
        ProgCaptionV2 Me, "Form Manipulation Example"
            ShowUserERR = True
End Sub

Private Sub Form_Resize()
' Move the controls as frmMain is resized
Frame1.Left = (frmMain.width - Frame1.width) / 2
Label1.Left = (frmMain.width - Label1.width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMain = Nothing
End
End Sub

Private Sub mnuFileExit_Click()
Unload Me
Set frmMain = Nothing
End
End Sub

Private Sub mnuFileRollDown_Click()
mnuView.Visible = False
'DynamicRollDown Me
BooleanRollup Me, False
End Sub

Private Sub mnuHelpAbout_Click()
'MsgBox "Simple Form Manipulation Routines. !Updated 1/26/2000!", vbOKOnly, "About"
CustomAbout "", "Simple Form Manipulation Routines. !Updated 2/10/2000", _
"Hey, if you use any of my code, please add me to your credits. " _
& "The source code for this project is given as is.", vAppName, True
'frmAd.Show
End Sub

Private Sub mnuHelpImprovements_Click()
MsgBox "If you have any improvements to this module, please send them to me." & Chr(10) & Chr(10) & "Thank you" & Chr(10) & "Michael Heath", vbOKOnly, "Send Me Improvements"
vEmailMe "Improvements_to_FormManipulator"
End Sub

Private Sub mnuHelpSend_Click()
MsgBox "Your default email client will now launch.  If you don't have an email client then Send your Questions or Comments to mheath@indy.net. Subject: Form Manipulation", vbOKOnly, "Email Me"
vEmailMe "Form_Manipulator"
End Sub

Private Sub mnuHelpUpdates_Click()
vQuestion = MsgBox("If you wish to be emailed when these subs are updated then click the Yes button, otherwise click the No button.", vbYesNo, "Send Update?")
Select Case vQuestion
    Case vbYes
        vEmailMe "Send_Me_Form_Manipulator_Updates"
    Case vbNo
End Select
End Sub

Private Sub mnuOptionsErrorOff_Click()
ShowUserERR = False
End Sub

Private Sub mnuOptionsErrorOn_Click()
ShowUserERR = True
End Sub

Private Sub tmrNoOffScreen_Timer()
NoOffScreen Me
NoOffScreen frmPopup
End Sub
