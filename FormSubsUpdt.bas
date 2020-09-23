Attribute VB_Name = "FormSubsUpdt"
' Module Created by Michael Heath
' Purpose of this module is to create a couple form tricks
' See each sub and read their descriptions
' Created 1/20/2000
' Updated 1/24/2000

' Set the Name of your application using Underscores so it will fit in
' subject line of email
' Following constants are used for naming the news and update text from the server.
' Rename them accordingly
Public Const vAppName = "Form_Manipulator"
Public Const strNewsFile = "news.txt"
Public Const strUpdateFile = "FormManupdt.txt"
Public Const strSetHost = "ftp://win2000"
Public Const vIniKey = "FormMan"
' End News and Updates

' End News and Updates

Public ShowUserERR As Boolean ' Used to decide if Error Message will popup or not
Public strPosTop As String ' Holds vForm.Top Value
Public strPosLeft As String ' Holds vForm.Left Value
' The Following values are used to set the height and width of a form
' They will be used in Form_Load, popupmenu RollUp & popupmenu RollDown
Public vHeight As String ' Holds vForm.Height Value
Public vWidth As String ' Holds vForm.Width Value
Public Const LongWidth = 8865
Public Const NormalWidth = 1275
Public Const RollUpTop = 600
Public Const RollUpLeft = 2000
Public Const NormalHeight = 7365
' Used to decide which direction you want to launch vForm from vMain
Public Const vRight = 0
Public Const vLeft = 1
Public Const vTop = 2
Public Const vBottom = 3
Public Const vTopCenter = 4
Public Const vBottomCenter = 5
Public Const vRightCenter = 6
Public Const vLeftCenter = 7
' Used in frmPosition
Public Const vCornerTopR = 0
Public Const vCornerTopL = 1
Public Const vCornerBtmL = 2
Public Const vCornerBtmR = 3
' Following declares are used for setting a form ontop or not ontop
 Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
' End

' Following Delcares are used for Killing the X button on a form
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&
' End Killing X button

Sub frmOnTop(vForm As Form, OnTop As Boolean)

' This sub combines the StayOnTop and NotOntop Subs

    ' Sub/Function Name       : OnTopFrm
    ' Purpose                 : Makes a form stay ontop or removes from ontop
    ' Parameters              : Form Object
    ' Created by              : Unknown
    ' Date Created            : Unkown
On Error GoTo LogERR
Dim width, height As Integer 'This will save the form's height and width cus we'll need it later

    Select Case OnTop
        Case True
            ' Place the form ontop
            width = vForm.width 'put the form's width
            height = vForm.height 'put the form's height
            SetWindowPos vForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, &H1 'Or &H2
            'Put form always On top, it will re-shape ur form
            frmCenterMe vForm
            
        Case False
            ' Disable the form from being ontop
            width = vForm.width 'put the form's width
            height = vForm.height 'put the form's height
            SetWindowPos vForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, &H1
            frmCenterMe vForm
    End Select
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "frmOnTop", Err.Description
End Sub
Public Sub frmCenterMe(vForm As Form)
    ' Sub/Function Name       : frmCenterMe
    ' Purpose                 : Centers a form on the screen
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
On Error GoTo LogERR
vForm.Left = (Screen.width - vForm.width) / 2
vForm.Top = (Screen.height - vForm.height) / 2
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "frmCenterMe", Err.Description

End Sub

Public Sub SaveFrmDiminsions(vForm As Form)
    ' Sub/Function Name       : SaveFrmDiminsions
    ' Purpose                 : Saves a forms diminsions to a file... ie (vForm.Width & vForm.height)
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : RestoreFrmDiminsions
On Error GoTo LogERR
' Save the current diminsions of vForm to a file.
' 2 Possible ways just as in SaveFrmPos, but we'll only show one for this one.
' 2:  Saves Everything to One file with Unique Keys
writeINI "FormHeight", vForm.Name, vForm.height, App.Path & "\" & "frmPos.psn"
    writeINI "FormWidth", vForm.Name, vForm.width, App.Path & "\" & "frmPos.psn"
' End
' The contents of this sub could also be combined with SaveFrmPos
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "SaveFrmDiminsions", Err.Description

End Sub

Public Sub RestoreFrmDiminsions(vForm As Form)
    ' Sub/Function Name       : RestoreFrmDiminsions
    ' Purpose                 : Saves a forms diminsions to a file... ie (vForm.Width & vForm.height)
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : SaveFrmDiminsions
' Read the Saved diminsions of vForm
' We could use either file we created just as with SaveFrmPos.
On Error GoTo LogERR
'Method Two:
    vHeight = ReadINI("FormHeight", vForm.Name, App.Path & "\" & "frmPos.psn")
        If vHeight = "" Then Exit Sub
    vWidth = ReadINI("FormWidth", vForm.Name, App.Path & "\" & "frmPos.psn")
        If vWidth = "" Then Exit Sub

' Common to both methods:
Dim intHeight As Double
    Dim intWidth As Double
        intHeight = vHeight
            intWidth = vWidth
    vForm.width = intWidth
vForm.height = intHeight
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "RestoreFrmDiminsions", Err.Description


End Sub

Public Sub SaveFrmPos(vForm As Form)
    ' Sub/Function Name       : SaveFrmPos
    ' Purpose                 : Saves a forms position to a file. (Two Methods)
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : ReSnapFrm
On Error GoTo LogERR
' Save the current position of vForm to a file by the name of the form.
' 2 Possible ways:
' 1:  This makes a new file for each Form and is unique in name like the form itself is
writeINI "Position", "FormTop", vForm.Top, App.Path & "\" & vForm.Name & ".psn"
    writeINI "Position", "FormLeft", vForm.Left, App.Path & "\" & vForm.Name & ".psn"
    
' 2:  Saves Everything to One file with Unique Keys
writeINI "FormTop", vForm.Name, vForm.Top, App.Path & "\" & "frmPos.psn"
    writeINI "FormLeft", vForm.Name, vForm.Left, App.Path & "\" & "frmPos.psn"
' End

' Troubleshooting Stuff
writeINI "TroubleShooter", "ScreenHeight", Screen.height, App.Path & "\" & vForm.Name & ".psn"
writeINI "TroubleShooter", "ScreenWidth", Screen.width, App.Path & "\" & vForm.Name & ".psn"
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "SaveFrmPos", Err.Description


End Sub

Public Sub ReSnapFrm(vForm As Form)
    ' Sub/Function Name       : ReSnapFrm
    ' Purpose                 : Repositions the form from info saved from SaveFrmPos. (Two Methods)
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : SaveFrmPos
On Error GoTo LogERR
' Read the Saved Position of vForm
' We could use either file we created.
' Method One:
    strPosTop = ReadINI("Position", "FormTop", App.Path & "\" & vForm.Name & ".psn")
        If strPosTop = "" Then Exit Sub
    strPosLeft = ReadINI("Position", "FormLeft", App.Path & "\" & vForm.Name & ".psn")
        If strPosLeft = "" Then Exit Sub
        
'Method Two:
    strPosTop = ReadINI("FormTop", vForm.Name, App.Path & "\" & "frmPos.psn")
        If strPosTop = "" Then Exit Sub
    strPosLeft = ReadINI("FormLeft", vForm.Name, App.Path & "\" & "frmPos.psn")
        If strPosLeft = "" Then Exit Sub

' Common to both methods:
Dim intPosTop As Double
    Dim intPosLeft As Double
        intPosTop = strPosTop
            intPosLeft = strPosLeft
    vForm.Top = intPosTop
vForm.Left = intPosLeft
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "ReSnapFrm", Err.Description

End Sub

Public Sub LoadNextToMain(vForm As Form, vMain As Form, vDirection As Integer)
    ' Sub/Function Name       : LoadNextToMain
    ' Purpose                 : Loads a form to the right, left, top or bottom of another form
    ' Parameters              : Form Object to load, Form Object to load next to
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : None
On Error GoTo LogERR
' vMain is the Main Form You choose
' vForm is the Form you want to load around VMain
' Get Posistion of vMain
Dim strMainTop As Long
Dim strMainLeft As Long
Dim strFormTop As Long
Dim strFormLeft As Long
Select Case vDirection
    Case 0
        'vRight
        strMainLeft = vMain.Left
        strFormLeft = strMainLeft + vMain.width
        vForm.Left = strFormLeft
        vForm.Top = vMain.Top
    Case 1
        'vLeft
        strMainLeft = vMain.Left
        strFormLeft = strMainLeft - vForm.width
        vForm.Left = strFormLeft
        vForm.Top = vMain.Top
    Case 2
        'vTop
        strMainTop = vMain.Top
        strFormTop = strMainTop - vForm.height
        vForm.Top = strFormTop
        vForm.Left = vMain.Left
    Case 3
        'vBottom
        strMainTop = vMain.Top
        strFormTop = strMainTop + vMain.height
        vForm.Top = strFormTop
        vForm.Left = vMain.Left
    Case 4
        'vTopCenter
        strMainTop = vMain.Top
        strFormTop = strMainTop - vForm.height
        vForm.Top = strFormTop
        vForm.Left = vMain.Left + ((vMain.width - vForm.width) / 2)
    Case 5
        'vBottomCenter
        strMainTop = vMain.Top
        strFormTop = strMainTop + vMain.height
        vForm.Top = strFormTop
        vForm.Left = vMain.Left + ((vMain.width - vForm.width) / 2)
    Case 6
        'vRightCenter
        strMainLeft = vMain.Left
        strFormLeft = strMainLeft + vMain.width
        vForm.Left = strFormLeft
        vForm.Top = vMain.Top + (vMain.height - vForm.height) / 2
    Case 7
        'vLeftCenter
        strMainLeft = vMain.Left
        strFormLeft = strMainLeft - vForm.width
        vForm.Left = strFormLeft
        vForm.Top = vMain.Top + (vMain.height - vForm.height) / 2
        
End Select
vForm.Show
' Make Sure this action doesn't move vForm off screen
NoOffScreen vForm
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name & vMain.Name, "LoadNextToMain", Err.Description

End Sub

Public Sub NoOffScreen(vForm As Form)
    ' Sub/Function Name       : NoOffScreen
    ' Purpose                 : Prevents a form from being placed offscreen if run inside a timer
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
    ' Sister Sub/Function     : SaveFrmPos
On Error GoTo LogERR
' Neat Little Corner Snap - Prevents app from going off screen
If vForm.WindowState = 1 Then Exit Sub ' form cannot be moved if minimized
    If vForm.Top < Screen.height - Screen.height + 10 Then ' Top Less than 10
         vForm.Top = Screen.height - Screen.height + 10
    End If
            If vForm.Left < Screen.width - Screen.width + 10 Then ' Left Less than 10
            vForm.Left = Screen.width - Screen.width + 10
    End If
         If vForm.Top > Screen.height - vForm.height - 10 Then ' Left more than 10
             vForm.Top = Screen.height - vForm.height - 10
    End If
            If vForm.Left > Screen.width - vForm.width - 10 Then ' Bottom more than 10
            vForm.Left = Screen.width - vForm.width - 10
    End If
        Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "NoOffScreen", Err.Description

End Sub
Public Sub BooleanRollup(vForm As Form, RollUp As Boolean)
    ' Sub/Function Name       : BooleanRollup
    ' Purpose                 : Window blind your form.
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000

On Error GoTo LogERR
' This sub combines the old subs DynamicRollUp and DynamicRollDown from my first version
Dim strHeight As String
Dim strWidth As String
Dim intDivHeight As Long

    Select Case RollUp
        Case True
        ' Shrinks a form in a cool little window blind way
            ' Roll the form up
            SaveFrmDiminsions vForm
                If vForm.height = 600 Then Exit Sub
            SaveFrmPos vForm
                intDivHeight = vForm.height / 100

                For x = 1 To 100
                    vForm.height = vForm.height - intDivHeight
                    Debug.Print vForm.height
                    vForm.Refresh
                Next x
                    vForm.height = RollUpTop

        Case False
        ' Returns a form from its shrunkin position
            ' Roll the form down
            strHeight = ReadINI("FormHeight", vForm.Name, App.Path & "\" & "frmPos.psn")
                If strHeight = "" Then Exit Sub
            intHeightCvt = strHeight
                If vForm.height = intHeightCvt Then Exit Sub
                intDivHeight = intHeightCvt / 100
                    For x = 1 To 100
                        vForm.height = vForm.height + intDivHeight
                        vForm.Refresh
                    Next x
                        vForm.height = intHeightCvt

    End Select
        Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "BooleanRollUp", Err.Description

End Sub


Public Sub frmPosition(vForm As Form, vPosition As Integer)
    ' Sub/Function Name       : frmPosition
    ' Purpose                 : Sets a form in one of the 4 corners of screen
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
On Error GoTo LogERR
Select Case vPosition
    Case 0
    ' TopRight
        vForm.Top = Screen.height - Screen.height
            vForm.Left = Screen.width - vForm.width
    Case 1
    ' TopLeft
        vForm.Top = Screen.height - Screen.height
            vForm.Left = Screen.width - Screen.width
    Case 2
    ' BottomLeft
        vForm.Top = Screen.height - vForm.height
            vForm.Left = Screen.width - Screen.width
    Case 3
    ' BottomRight
        vForm.Top = Screen.height - vForm.height
            vForm.Left = Screen.width - vForm.width
End Select
        vForm.Show
        NoOffScreen vForm
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "frmPosition", Err.Description

End Sub

Public Sub ProgCaption(vForm As Form, Message As String)
    ' Sub/Function Name       : ProgCaption
    ' Purpose                 : Sets a form's caption
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
On Error GoTo LogERR
If Message = "" Then
    vForm.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
Else
    vForm.Caption = App.EXEName & " - " & Message
End If
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "ProgCaption", Err.Description

End Sub
Public Sub ProgCaptionV2(vForm As Form, Message As String)
    ' Sub/Function Name       : ProgCaption
    ' Purpose                 : Sets a form's caption
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/20/2000
On Error GoTo LogERR
' Set the following string to be the Name of YOUR app.
If Message = "" Then
    vForm.Caption = vAppName & " v" & App.Major & "." & App.Minor & "." & App.Revision
Else
    vForm.Caption = vAppName & " - " & Message
End If
    Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "ProgCaptionV2", Err.Description

End Sub
Public Sub frmKillExit(vForm As Form)
    ' Sub/Function Name       : frmKillExit
    ' Purpose                 : Disables the X button on a form
    ' Parameters              : Form Object
    ' Created by              : Unknown
    ' Date Created            : Unknown
On Error GoTo LogERR
    Dim hSysMenu As Long, nCnt As Long
    ' Get handle to our form's system menu
    ' (Restore, Maximize, Move, close etc.)
    hSysMenu = GetSystemMenu(vForm.hwnd, False)

    If hSysMenu Then
        ' Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            ' Menu count is based on 0 (0, 1, 2, 3...)
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE ' Remove the seperator
            DrawMenuBar vForm.hwnd
            ' Force caption bar's refresh. Disabling X button
        End If
    End If
        Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "frmKillExit", Err.Description

End Sub

Public Sub CustomMsg(Message As String, vCaption As String, vIcon As Image)
    ' Sub/Function Name       : CustomMsg
    ' Purpose                 : Creates a Message Box out of a form without using API
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000
' This sub was just created for the hell of it. No big deal and not much
' to it. Not a lot of detail in it either.  To add more detail, like a cancel
' button, yes, no, retry, etc... then just add them to the form and make the
' visible property=false.  Then whatever button you need, just make visible = true
' in this sub.  You will also need to add the strings above.  You could also go into
' further detail and make the lblmsg.width and height dynamic in this sub.
' This is only a sample sub and no sample of creating the message box has been
' provided.  If you find that you need an example of this routine, please email me
' at mheath@indy.net and ask for a sample project using this.

' Things Needed:
'   1. Form, name it  frmMsgBox
'   2. Label on form. Name it lblMsg
'   3. Image Box. Name it imgIcon
'   4. Command button, name it cmdOK
'   5. Icons.  VB5 and 6 come with icons in the graphics directory (in an example I used
'      an additional form and named it frmIcons.  I put all the icons on that form
'      and just called them when I needed them.)
    
On Error GoTo LogERR
frmMsgBox.Caption = vCaption
frmMsgBox.Icon = vIcon.Picture
    frmMsgBox.lblMsg.Caption = Message

        frmMsgBox.height = (frmMsgBox.lblMsg.height + frmMsgBox.cmdOk.height) + (frmMsgBox.cmdOk.height * 5)
        frmMsgBox.imgIcon.Picture = vIcon.Picture
        frmMsgBox.imgIcon.Top = ((frmMsgBox.height - frmMsgBox.height) + frmMsgBox.imgIcon.height)
        frmMsgBox.lblMsg.Left = (frmMsgBox.imgIcon.width + frmMsgBox.imgIcon.Left)
        frmMsgBox.lblMsg.Top = frmMsgBox.imgIcon.Top '- frmMsgBox.cmdOK.height
           frmMsgBox.cmdOk.Left = (frmMsgBox.width - frmMsgBox.cmdOk.width) / 2
                frmMsgBox.cmdOk.Top = frmMsgBox.height - (frmMsgBox.cmdOk.height * 3)
                frmMsgBox.Show
                   frmCenterMe frmMsgBox
      Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String

    LogMyErrors "frmMsgBox", "CustomMsg", Err.Description
      
End Sub

Public Sub vEmailMe(Message As String)
' Change the email addy to yours and the subject to whatever you need
Dim vShell As String
vShell = "Start.exe mailto:mheath@indy.net?Subject="
vShell = vShell & Message
nResult = Shell(vShell, vbHide)
End Sub

Public Sub StreakExit(vForm As Form)
    ' Sub/Function Name       : StreakExit
    ' Purpose                 : Moves a form across the top and down the side on exit
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000
On Error GoTo LogERR
Dim vScreenWidth As Long
Dim vScreenHeight As Long
frmPosition vForm, vCornerTopL
    vScreenWidth = (Screen.width - vForm.width) / 100
        For x = 1 To 100
            vForm.Left = vForm.Left + vScreenWidth
            NoOffScreen vForm
            vForm.Refresh
            Next x
            BooleanRollup vForm, True
            frmMain.mnuView.Visible = True
 vScreenHeight = (Screen.height - vForm.height) / 100
    For x = 1 To 100
        vForm.Top = vForm.Top + vScreenHeight
        'NoOffScreen vForm
        vForm.Refresh
    Next x
      Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "StreakExit", Err.Description
  
End Sub

Public Sub LogMyErrors(ErrorForm As String, ErrorSub As String, vError As String)
    ' Sub/Function Name       : LogMyErrors
    ' Purpose                 : Logs Any Errors to a file
    ' Parameters              : Form Object, Sub, Err.description
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000

' This sub Logs all errors created by these subs to the application path, file ErrorLog.txt
If vError = "" Then vError = "Unknown Error Occurred"
writeINI Date & "-" & Time & " " & ErrorForm, ErrorSub, vError, App.Path & "\ErrorLog.txt"
    If ShowUserERR = True Then
        MsgBox "The following error has been logged to " & App.Path & "\ErrorLog.txt" _
        & Chr(10) & vError & Chr(10) & Chr(10) & _
        "If the error continues, please send the ErrorLog.txt to the email address under the help menu." & Chr(10) & Chr(10) & "This message can be disabled under the options menu.", vbOKOnly, ErrorForm & " - Error"
    End If
End Sub

Public Sub frmShrink(vForm As Form)
    ' Sub/Function Name       : frmShrink
    ' Purpose                 : Shrinks a form while closing it
    ' Parameters              : Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000

Dim strHeight As String
Dim strWidth As String
Dim intDivHeight As Long
Dim intDivWidth As Long
On Error GoTo LogERR
     
        ' Shrinks a form in a cool little window blind way
            ' Roll the form up
            SaveFrmDiminsions vForm
                If vForm.height = 600 Then Exit Sub
            SaveFrmPos vForm
                intDivHeight = vForm.height / 100
                intDivWidth = vForm.width / 100
                For x = 1 To 100
                    vForm.height = vForm.height - intDivHeight
                    vForm.width = vForm.width - intDivWidth
                    Debug.Print vForm.height
                    vForm.Refresh
                Next x
                    vForm.height = RollUpTop
Unload vForm
Exit Sub
LogERR:
    ' Lets get the error and log it to a file in the application directory called
    ' ErrorLog.txt.  Now when your customer gets an error in your app, they can send
    ' you the errorlog.txt and you can do your further troubleshooting.
    ' ErrorForm As String, ErrorSub As String, vError As String
   
    LogMyErrors vForm.Name, "frmShrink", Err.Description
  
End Sub
 Public Sub WriteSettings()
    ' Sub/Function Name       : WriteSettings
    ' Purpose                 : Saves some settings to a file for the purpose of sending
    '                           in a request for update via FTP
    ' Parameters              : news.txt, updates.ini, programnameupdt.txt, new exe file download
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000

 
 ' Write down the settings for autoupdating the program
writeINI "Update", "Revision", App.Revision, App.Path & "\settings.ini"
writeINI "Files", "News", strNewsFile, App.Path & "\settings.ini"
writeINI "Files", "Update", strUpdateFile, App.Path & "\settings.ini"
writeINI "Remote", "IP", strSetHost, App.Path & "\settings.ini"
writeINI "Files", "Key", vIniKey, App.Path & "\settings.ini"

'"209.183.122.111", App.Path & "\settings.ini"
 End Sub
Public Sub CenterSnglFrame(vFrame As Frame, vForm As Form, HeightorWidth As Integer)
    ' Sub/Function Name       : CenterSnglFrame
    ' Purpose                 : Centers "ONE" frame on a form. Place all your controls
    '                         : on a frame and then center it
    ' Parameters              : Frame Object, Form Object
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000

Select Case HeightorWidth
    Case 1
        ' Center Height
        vFrame.Top = (vForm.height - vFrame.height) / 2
    Case 2
        ' Center Width
         vFrame.Left = (vForm.width - vFrame.width) / 2
    Case 3
        ' Center Both
         vFrame.Top = (vForm.height - vFrame.height) / 2
         vFrame.Left = (vForm.width - vFrame.width) / 2
End Select
End Sub
Public Sub CustomAbout(vVersion As String, vDescript As String, vDisclaim As String, vCaption As String, v3D As Boolean)
    ' Sub/Function Name       : CustomAbout
    ' Purpose                 : Create your own custom About box
    ' Parameters              : Form Object, AppVersion, AppDescription, AppDisclaimer
    '                         : Caption String, Form Appearance
    ' Created by              : Michael Heath
    ' Date Created            : 1/22/2000
' This routine requires that the frmAout.frm be placed in the project
' I know it's a little simple, but this is a beginner module.
' The CustomAbout Sub is in no way a really awesome little routine. It has a lot
' of room for improvements.  If you have any, please share them.
If v3D = True Then
    frmAbout.Appearance = 1
Else
    frmAbout.Appearance = 0
End If
    frmAbout.lblCaption.Caption = "About - " & vCaption
        frmAbout.lblDescription.Caption = vDescript
            frmAbout.lblDisclaimer.Caption = vDisclaim
If vVersion = "" Then
    frmAbout.lblVersion.Caption = vAppName & " v" & App.Major & "." & App.Minor & "." & App.Revision
Else
    frmAbout.lblVersion.Caption = vVersion
End If
    frmAbout.frDescription.BackColor = frmAbout.BackColor
        frmAbout.frDisclaimer.BackColor = frmAbout.BackColor
            frmAbout.cmdOk.BackColor = frmAbout.BackColor
Load frmAbout
frmAbout.Show

End Sub
Public Sub Main()
' This is our startup sub.  You could also move the WriteSettings to the Main Form
' And start with it instead of this sub
WriteSettings
    Load frmMain
        frmMain.Show
End Sub
