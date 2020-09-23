Attribute VB_Name = "ReadWriteIni"
 Option Explicit
'<-- Used for Functions ReadINI & WriteINI
 #If Win16 Then

   Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
   Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal Default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
   
   #Else
    
   Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
   #End If
'<------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->
Public ININum As Integer
Public GrabUrl(1 To 365) As String
Public x As Integer
Public strHits As String
Public Hits As Double

    ' Sub/Function Name       : ReadINI
    ' Purpose                 : Reads info from an INI file
    ' Parameters              : Strings
    ' Created by              : Unknown
    ' Date Created            : Unkown
Function ReadINI(Section, KeyName, filename As String) As String
       
       Dim sRet As String
       sRet = String(255, Chr(0))
       ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
   
   End Function
    
    ' Sub/Function Name       : writeINI
    ' Purpose                 : Writes info to an INI file
    ' Parameters              : Strings
    ' Created by              : Unknown
    ' Date Created            : Unkown

Function writeINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
       
       Dim R
       R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
   
End Function


