Attribute VB_Name = "Main"
Option Explicit

'Declare Global function
'Get color of active window
 Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long

'Declare global constant
'Color of active window
 Public Const COLOR_ACTIVECAPTION = 2
 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public gwsWorkspace As Workspace
Global sRutaIni     As String ' ruta de archivo ini del sistema
Global sRutaBase    As String
Global sruta As String
Public Function LeerIni(seccion As String, clave As String) As String
    Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetPrivateProfileString(seccion, clave, "", retVal, Len(retVal), sRutaIni)
    
    If worked = 0 Then
        LeerIni = ""
    Else
        LeerIni = Left(retVal, InStr(retVal, Chr(0)) - 1)
    End If
End Function

Public Function GrabarIni(seccion As String, clave As String, cValor As String)
    Dim retVal As String, AppName As String, worked As Integer
    worked = WritePrivateProfileString(seccion, clave, cValor, sRutaIni)
End Function



