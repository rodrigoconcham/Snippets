VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Autor del módulo cSuperClass : Paul Caton ( para usar los CallBacks dentro del cls)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Módulo con funciones básicas para listar ventanas y procesos en vb
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Autor : Luciano L.E - www.recursosvisualbasic.com.ar

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private z_IDEflag           As Long         'Flag indicating we are in IDE
Private z_ScMem             As Long         'Thunk base address
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_CALLBACKORDINAL As Long = 22 ' Ubound(callback thunkdata)+1, index of the callback

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''
'Udt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Enum EWindowState
    [eNormal] = 1
    [eMinimized] = 2
    [eMaximized] = 3
End Enum


'Para los iconos de archivos y procesos
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

 
'Declaraciones Api para las ventanas
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindowx Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'--------Constants used with APIs
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
Private Const GW_OWNER = 4
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_APPWINDOW = &H40000




Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


' procesos
Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Type WTS_PROCESS_INFO
   SessionID As Long
   ProcessID As Long
   pProcessName As Long
   pUserSid As Long
End Type


Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)

Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&

' timer
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


' terminar procesos
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type


' eventos
''''''''''''''''''''''''''''''

' envia información de las ventanas/aplicaciones
Event GetWindows( _
    ByVal lHwnd As Long, _
    ByVal sTitle As String, _
    ByVal classWindowName As String, _
    ByVal WindowState As EWindowState, _
    ByVal wLeft As Long, _
    ByVal wTop As Long, _
    ByVal wRight As Long, _
    ByVal wBottom As Long)
    
Event GetWindowsChild( _
    ByVal lHwnd As Long, _
    ByVal lHwndParent As Long, _
    ByVal sTitle As String, _
    ByVal IsVisible As Boolean, _
    ByVal classWindowName As String, _
    ByVal WindowState As EWindowState, _
    ByVal wLeft As Long, _
    ByVal wTop As Long, _
    ByVal wRight As Long, _
    ByVal wBottom As Long)
    
    
' envia información de los procesos
Event GetProcess( _
    ByVal sNameProcess As String, _
    ByVal SpathProcess As String, _
    ByVal HandleProcess As Long)
    
' eventos que se disparan al cerrar ventanas y procesos pero desde el programa, no desde windows
Event CloseWindows( _
    ByVal sNameWindow As String, _
    lHwnd As Long, _
    sClassName As String)
    
Event CloseProcess( _
    ByVal sNameProcess As String, _
    ByVal SpathProcess As String, _
    ByVal HandleProcess As Long)
    
' Envia información de error
Event Error(ByVal sError As String)
' alcomienzo de la enumaración
Event StartEnumWindow()
Event StartEnumWindowChild()

Event StartEnumProcess()
' al Cambiar la lista de ventanas y procesos. Funciona cuando se establece el método Notify en true
Event ChangeWindowList()
Event ChangeProcessList()

Private mClassName                      As String
Private mHwndOwnerForm                  As Long
Private mProcessCount                   As Long
Private mWindowsVisibleCount            As Long
Private mbRaiseEventEnumWindows         As Boolean

' variables para la cantidad de ventanas
Private mCountNW                        As Long
Private iWl                             As Long
    
' variable para almacenar la cantidad de procesos
Private mCountLP                        As Long

'Private tWindowsList
'    sCaption As String
'    lHwnd    As Long
'End Type

'Private mArrWindowsList()               As tWindowsList
'Private mArrCurrentWindowsList()        As tWindowsList

''''''''''''''''''''''''''''''''''''''''''''''''''

'------------------------**********


'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
       Optional ByVal nOrdinal As Long = 1, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0
    Const INDX_CALLBACK As Long = 1
    Const INDX_EBMODE   As Long = 2
    Const INDX_BADPTR   As Long = 3
    Const INDX_EBX      As Long = 5
    Const INDX_PARAMS   As Long = 12
    Const INDX_PARAMLEN As Long = 17

    Dim z_Cb()    As Long    'Callback thunk array
    Dim nCallback As Long
      
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + 16  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError  ' oops
        Exit Function
    End If
    z_cbFunk.Add z_ScMem, "h" & nOrdinal                  'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(4) = &HBB60E089: z_Cb(6) = &H73FFC589: z_Cb(7) = &HC53FF04: z_Cb(8) = &H7B831F75: z_Cb(9) = &H20750008: z_Cb(10) = &HE883E889: z_Cb(11) = &HB9905004: z_Cb(13) = &H74FF06E3: z_Cb(14) = &HFAE2008D: z_Cb(15) = &H53FF33FF: z_Cb(16) = &HC2906104: z_Cb(18) = &H830853FF: z_Cb(19) = &HD87401F8: z_Cb(20) = &H4589C031: z_Cb(21) = &HEAEBFC
    
    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                  'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                       'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                  'Cache ordinal used for zTerminateThunks
      
    Debug.Assert zInIDE
    If bIdeSafety = True And z_IDEflag = 1 Then             'If the user wants IDE protection
        z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False)  'EbMode Address
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    z_Cb(INDX_PARAMLEN) = nParamCount * 4                   'Set the number of stck bytes to release on thunk return
      
    '\\LaVolpe - redirect address to proper location in virtual memory. Was: z_Cb(INDX_EBX) = VarPtr(z_Cb(INDX_OWNER))
    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer
      
    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + 16                       'Thunk code start address
    
End Function


'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
    ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
    If asUnicode Then
        zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
    Else
        zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
    End If
    Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
    ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
    Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                         'Address of the vTable
    Dim i     As Long                                                         'Loop index
    Dim j     As Long                                                         'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            ' \\LaVolpe - Added propertypage offset
            If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
                If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
                    Exit Function                                                   'Bail...
                End If
            End If
        End If
    End If
  
    i = i + 4                                                                 'Bump to the next entry
    j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If
    
        i = i + 4                                                               'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries
    Do While nAddr < nLimit                                                   'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
        If nEntry <> 0 Then                                                     'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Do                                                             'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop
End Function

Private Function zInIDE() As Long
    ' This is only run in IDE; it is never run when compiled
    z_IDEflag = 1
    zInIDE = z_IDEflag
End Function


Private Sub Class_Initialize()
    mbRaiseEventEnumWindows = True
End Sub

Sub TopMost(lHwnd As Long, bValue As Boolean)
   ' SetWindowPos lHwnd, Abs(bValue) - 2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub ShowWindows(lHwnd As Long, bValue As Boolean)
    ShowWindow lHwnd, Abs(bValue)
End Sub

Function SetActiveWindows(lHwnd As Long)
    
    SetForegroundWindow lHwnd
    If IsIconic(lHwnd) Then
       Me.ShowWindows lHwnd, True
    End If
    
End Function

' Para Buscar ventanas a partir del caption o el className y retornar el Handle
Function FindWindowByCaption(sCaption As String) As Long
    FindWindowByCaption = FindWindowx(vbNullString, Trim(sCaption))
End Function

Function FindWindowByClassName(sClassName As String) As Long
    FindWindowByClassName = FindWindowx(Trim(sClassName), vbNullString)
End Function


' Cerrar los timers al finalizar la clase
Private Sub Class_Terminate()
    On Error GoTo error_Function
    
    If mHwndOwnerForm <> 0 Then ' si se usó el método Notify , mHwndOwnerForm tiene el handle del form
        KillTimer mHwndOwnerForm, ObjPtr(Me) + 3 ' Timer de ventanas
        KillTimer mHwndOwnerForm, ObjPtr(Me) + 4 ' Timer de procesos
    End If
Exit Sub
error_Function:
RaiseEvent Error("Error en el class Terminate")
End Sub

' Retorna la cantidad de procesos abiertos
Property Get CountProcess() As Long
    
    Dim arrProc(250) As Long
    Dim i As Long
    
    EnumProcesses arrProc(1), 1000, i
    mProcessCount = i / 4
    CountProcess = mProcessCount
    DoEvents

End Property

' Retorna la cantidad de ventanas de la barra de tareas
Property Get CountWindowsVisible() As Long
    
    mbRaiseEventEnumWindows = False ' Flag para no lanzar el evento
    mWindowsVisibleCount = 0
    
    EnumWindows scb_SetCallbackAddr(2, 1), 1 ' establecer el CallBack ( Api EnumWindow - Función n° 1)
    CountWindowsVisible = mWindowsVisibleCount
    mbRaiseEventEnumWindows = True
End Property

' dibujar en un Hdc el ícono asociado a un tipo de fichero
Function DrawIconFileInDC( _
    sPathFile As String, _
    lHdc As Long, Optional X As Long = 0, _
    Optional Y As Long = 0, _
    Optional lWidth As Long = 32, _
    Optional lHeight As Long = 32) As Boolean
    
    On Error GoTo error_Function
    
    Dim hIcon As Long
    
    If Len(Dir(sPathFile)) = 0 Then
       RaiseEvent Error("No se encontró el archivo: " & sPathFile)
       Exit Function
    End If
    
    hIcon = ExtractAssociatedIcon(App.hInstance, sPathFile, 0)
    
    If hIcon = 0 Then
       RaiseEvent Error("Error al extraer el ícono del archivo: " & sPathFile)
    End If
    
    DrawIconFileInDC = DrawIconEx(lHdc, X, Y, hIcon, lWidth, lHeight, 0, 0, DI_NORMAL)

    DestroyIcon hIcon
    
Exit Function
error_Function:
RaiseEvent Error(Err.Description)
End Function

' Dibujar en un Hdc el ícono del caption de la ventana
Function DrawIconCaptionInDC( _
    lHwnd As Long, _
    lHdc As Long) As Boolean
    
    Dim ret As Long
    
    Dim r As RECT
    'ret = SetRect(r, 5, 7, 10, 10)
    ret = SetRect(r, 0, 0, 16, 16)
    DrawIconCaptionInDC = DrawCaption(lHwnd, lHdc, r, 20)
    
End Function

' Devolver el caption de una ventana con GetWindowText
Function GetWindowCaption(ByVal lHwnd As Long) As String

    Dim stemp As String
    Dim buffer As String * 256
    
    stemp = Space$(GetWindowTextLength(lHwnd) + 1)
    
    GetWindowText lHwnd, stemp, Len(stemp)
    
    stemp = Left$(stemp, Len(stemp) - 1)
    
    If LCase(stemp) <> "program manager" Then
       GetWindowCaption = stemp
    End If
    
End Function

' Cerrar una ventana a partir del Hwnd con SendMessage
Function CloseWindowByHandle(ByVal lHwnd As Long) As Long
    
    ' comproibar si la ventana existe
    If IsWindow(lHwnd) <> 1 Then
       RaiseEvent Error("No se encontró la ventana con Handle: " & lHwnd)
       Exit Function
    End If
    
    Dim lHandle As Long
    Dim strCaption As String
    Dim sClassName As String
    
    ' almacenar los datos de la ventana
    strCaption = Me.GetWindowCaption(lHwnd)
    sClassName = Me.GetClassWindowName(lHwnd)
    
    lHandle = lHwnd
    
    ' cerrarla
    CloseWindowByHandle = SendMessage(lHwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
    
    
    Sleep 200 ' no se por que pero si no, no lanza el evento
    
    ' El handle debe ser 0, si no hubo un error al cerrala
    If IsWindow(lHandle) = 0 Then
       RaiseEvent CloseWindows(strCaption, lHandle, sClassName) ' Ok
    Else
       RaiseEvent Error("Error al cerrar la ventana con Handle: " & lHwnd & " - " & strCaption) ' error
    End If
    
End Function

' Devolver el Hwnd de la ventana que tiene el  foco con el Api GetForegroundWindow
Function GetHwndActiveWindow()
    GetHwndActiveWindow = GetForegroundWindow()
End Function

' Devolver el nombre de clase de una ventana con getClassName
Function GetClassWindowName(lHwnd As Long) As String
    Dim buffer As String * 256
    Dim ret As Long
    ret = GetClassName(lHwnd, buffer, 256)
    GetClassWindowName = Left(buffer, ret)
End Function


' función privada para Enumerar los procesos abiertos
Private Sub mEnumerateProcesses(Optional bRaiseEvent As Boolean = True)
    Dim Array_Procesos() As Long
    Dim buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim sFileName As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
    
On Error GoTo error_Function
    
    ReDim Array_Procesos(250) As Long
    
    ' Obtiene un array con los id de los procesos
    ret = EnumProcesses(Array_Procesos(1), 1000, t_cbNeeded)

    i_Procesos = t_cbNeeded / 4
    
    ' Recorre todos los procesos
    For i = 1 To i_Procesos
    
            ' Lo abre y devuelve el handle
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
            
            If Handle_Proceso <> 0 Then
                ' Crea un buffer para almacenar el nombre y ruta
                buffer = Space(255)
                ' Le pasa el Buffer al Api y el Handle
                ret = GetModuleFileNameExA(Handle_Proceso, 0, buffer, 255)
                ' Le elimina los espacios nulos a la cadena devuelta
                Ruta = Left(buffer, ret)
                sFileName = GetFileName(Ruta)
            
            End If
            ' Cierra el proceso abierto
            ret = CloseHandle(Handle_Proceso)
                               
            If Array_Procesos(i) = 0 Or Array_Procesos(i) = 4 Then
                sFileName = "System"
                Ruta = ""
            End If
            ' cuando se llama desde la clase no lanzar el evento, si no notificar
            If bRaiseEvent Then
                RaiseEvent GetProcess(sFileName, Ruta, Array_Procesos(i))
            End If
            DoEvents
    Next
    
    'cantidad de procesos
    mProcessCount = i_Procesos + 1
    
Exit Sub
error_Function:
RaiseEvent Error(Err.Description)

End Sub

' devuelve True si el handle existe
Function WindowsExist(lHwnd As Long) As Boolean
    WindowsExist = CBool(IsWindow(lHwnd))
End Function

' recuperar  el nombre de archivo del Path
Private Function GetFileName(Spath As String) As String
    If Spath <> "" Then
        GetFileName = Right(Spath, Len(Spath) - InStrRev(Spath, "\"))
    End If
End Function

' función pública para cerrar un proceso por el handle
Function CloseProcessByHandle(lProcessID As Long) As Boolean
    CloseProcessByHandle = mCloseProcessByHandle(lProcessID)
End Function

' Función del sitio web : http://www.visualbasic.happycodings.com/Windows_and_Controls/code44.html
' Matar un proceso por el handle
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function mCloseProcessByHandle( _
    lProcessID As Long, _
    Optional bRaiseEventError As Boolean = True) As Boolean

On Error GoTo error_Function
    
    Dim lhwndProcess As Long
    Dim lExitCode As Long
    Dim lRetVal As Long
    Dim lhThisProc As Long
    Dim lhTokenHandle As Long
    Dim tLuid As LUID
    Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINATE = &H1
    Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Const SE_PRIVILEGE_ENABLED = &H2

    On Error Resume Next
    
    'If lHwndWindow Then
        'Get the process ID from the window handle
    '    lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
    'End If
    
    If lProcessID Then
        'Give Kill permissions to this process
        lhThisProc = GetCurrentProcess
        
        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        'Set the number of privileges to be change
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        'Enable the kill privilege in the access token of this process
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

        'Open the process to kill
        lhwndProcess = OpenProcess(PROCESS_TERMINATE, 0, lProcessID)
    
        If lhwndProcess Then
            'Obtained process handle, kill the process
            lRetVal = CBool(TerminateProcess(lhwndProcess, lExitCode))
            Call CloseHandle(lhwndProcess)
            
            If bRaiseEventError Then
                If lRetVal = 0 Then
                    RaiseEvent Error("Error al intentar cerrar el proceso: " & CStr(lProcessID))
                Else
                    RaiseEvent CloseProcess("", "", lProcessID)
                End If
            End If
            
            mCloseProcessByHandle = lRetVal
        ElseIf bRaiseEventError Then
            RaiseEvent Error("No se encontró el proceso o no tiene permisos para cerrarlo:  Proceso con handle: " & lProcessID)
        End If
    End If
    On Error GoTo 0
    
Exit Function
error_Function:
    RaiseEvent Error("Error al intentar cerrar el proceso: " & CStr(lProcessID))
    
End Function

' Cerrar un proceso por el nombre
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CloseProcessByName(sNameProcess As String) As Boolean
    
    On Error GoTo error_Function

    Dim arrProc() As Long, buffer As String, i As Long, ret As Long
    Dim Ruta As String, sFile As String, t_cbNeeded As Long, hProc As Long, iCountProc As Long
    
    Dim bFind As Boolean
    
    ReDim arrProc(250) As Long
    
    ret = EnumProcesses(arrProc(1), 1000, t_cbNeeded)

    iCountProc = t_cbNeeded / 4
    ' recorrer los procesos
    For i = 1 To iCountProc
    
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION + PROCESS_VM_READ, 0, arrProc(i))
            
            If hProc <> 0 Then
                buffer = Space(255) ' pBuffer para el path
                ret = GetModuleFileNameExA(hProc, 0, buffer, 255)
                Ruta = Left(buffer, ret)
                sFile = GetFileName(Ruta) ' Obtener el nombre del exe
                
                ' comparar si es el mismo
                If Trim(LCase(sFile)) = Trim(LCase(sNameProcess)) Then
                                   
                    ret = mCloseProcessByHandle(arrProc(i), False)
                    
                    If ret Then
                        RaiseEvent CloseProcess(sFile, Ruta, arrProc(i))
                    Else
                        ' error
                        RaiseEvent Error("Error al cerrar el proceso: " & CStr(arrProc(i) & " - " & sFile))
                    End If
                    
                    bFind = True
                    
                End If
            End If
            ret = CloseHandle(hProc)
    Next
    
    If bFind = False Then
       RaiseEvent Error("No se encontró el proceso: " & sNameProcess)
    End If
    
Exit Function
error_Function:
    RaiseEvent Error("No se encontró el proceso: " & sNameProcess)
    
End Function

' Cerrar todas las ventanas de la clase indicada
Function CloseWindowsByClassName(sNameClassName As String) As Boolean
    mClassName = sNameClassName ' variable local para usarla desde mCallBackCloseByClassName y comparar el nombre de clase si el mismo para poder cerrar la ventana
    Call EnumWindows(scb_SetCallbackAddr(2, 3), 1) ' Disparar el CallBack "mCallBackCloseByClassName"  - Función n° 3
    mClassName = ""
End Function

' Función pública para enumerar los procesos
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EnumerateProcesses()
    
    ' Comienzo de la enumeración de procesos
    RaiseEvent StartEnumProcess
    
    ' enumerar
    Call mEnumerateProcesses
    
End Sub


' función privada para Enumerar la lista de ventanas con EnumWindow
Private Sub mEnumWindow(handle As Long)
        
    ' Flag para lanzar o no el evento que envia los datos de las ventanas
    If mbRaiseEventEnumWindows Then
        
        If GetWindowsTaskBar(handle) <> "" Then
            Dim tWP As WINDOWPLACEMENT
    
            With tWP
                tWP.Length = Len(tWP)
        
                GetWindowPlacement handle, tWP
            
                'Disparar el evento y pasar los valores
                RaiseEvent GetWindows( _
                                    handle, _
                                    GetWindowCaption(handle), _
                                    GetClassWindowName(handle), _
                                    .showCmd, _
                                    .rcNormalPosition.Left, _
                                    .rcNormalPosition.Top, _
                                    .rcNormalPosition.Right, _
                                    .rcNormalPosition.Bottom)
            End With
        End If
    
    ElseIf GetWindowsTaskBar(handle) <> "" Then
        mWindowsVisibleCount = mWindowsVisibleCount + 1 ' cantidad de ventanas ( solo usada para cuando se llama desde la propiedad WindowsVisibleCount)
    End If

End Sub

' Función privada para Enumerar las ventanas Child ( esta función es usada desde CallBackEnumWindowChild - Función 2)
Private Sub mEnumWindowChild(handle As Long)
        
    Dim tWP As WINDOWPLACEMENT
    
    With tWP
        tWP.Length = Len(tWP)
        
        GetWindowPlacement handle, tWP
            
        'Disparar el evento y pasar los valores
        RaiseEvent GetWindowsChild( _
                              handle, _
                              GetParent(handle), _
                              GetWindowCaption(handle), _
                              IsWindowVisible(handle), _
                              GetClassWindowName(handle), _
                             .showCmd, _
                             .rcNormalPosition.Left, _
                             .rcNormalPosition.Top, _
                             .rcNormalPosition.Right, _
                             .rcNormalPosition.Bottom)
    End With
    
End Sub

' Función pública para enumerar las ventanas
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnumWindow()
    
    'Lanzar evento para notificar el Comienzo de la enumeración de ventanas
    RaiseEvent StartEnumWindow
    
    ' Enumerar todas las ventanas
    EnumWindows scb_SetCallbackAddr(2, 1), 1 ' establecer el CallBack de la función 1 que usa EnumWindow par alas ventanas de la barra de tareas
    
End Sub

' Función que devuelve las ventanas hijas
Public Sub EnumWindowChild(lHwnd As Long)
  
    'Lanzar evento para notificar el Comienzo de la enumeración de ventanas
    RaiseEvent StartEnumWindowChild
  
  ' Llamar a EnumChildWindows pasándole la 2° función CallBack : CallBackEnumWindowChild
   EnumChildWindows lHwnd, scb_SetCallbackAddr(2, 2), ByVal 0&

End Sub

' Función que inicializa o detiene el timer quer notifica la lista de ventanas de la barra de tareas
Function NotifyChangeWindow(frm As Form, bEnable As Boolean)

    Dim ret As Long
    
    mCountNW = 0
    iWl = 0
    
    mHwndOwnerForm = frm.hwnd
    
    ' Crear el Timer para poder Notificar
    If bEnable Then
    
        ' cerrar el timer por las dudas
        ret = KillTimer(mHwndOwnerForm, ObjPtr(Me) + 3)
        
        ' lanzar el timer
        Call SetTimer(mHwndOwnerForm, ObjPtr(Me) + 3, 50, scb_SetCallbackAddr(4, 4))
    
    ' ... cerrar el timer y finalizar la notificación
    Else
        ret = KillTimer(mHwndOwnerForm, ObjPtr(Me) + 3)
    End If
        
End Function

' Función que inicializa la notificación para saber cuando cambia la lista de Procesos
Function NotifyChangeProcess(frm As Form, bEnable As Boolean)

    Dim ret As Long
    
    mCountLP = 0
    mHwndOwnerForm = frm.hwnd
    
    ' Notificar
    If bEnable Then
        ' cerrar el timer por las dudas
        ret = KillTimer(mHwndOwnerForm, ObjPtr(Me) + 4) ' mata la función CallBackEnumProcessNotify
        
        ' lanzar el timer
        Call SetTimer(mHwndOwnerForm, ObjPtr(Me) + 4, 50, scb_SetCallbackAddr(4, 6))
    
    ' ... cerrar el timer de procesos y finalizar la notificación
    Else
        ret = KillTimer(mHwndOwnerForm, ObjPtr(Me) + 4) ' ' mata la función CallBackEnumProcessNotify
    End If
End Function


' Devuelve el caption de la ventana de la barra de tareas
Private Function GetWindowsTaskBar(hwnd As Long) As String

    Dim lReturn As Long
    Dim lExStyle As Long
    Dim bNoOwner As Boolean
    Dim sWindowText As String
    
    If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
       
       bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
       
       lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
       
       If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                
            sWindowText = Space$(256)
            lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                
            If lReturn Then
                sWindowText = Left$(sWindowText, lReturn)
                GetWindowsTaskBar = sWindowText
            End If
                
       End If
    End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Comienzo de Funciones CallBack : ' importante : No colocar código en este bloque

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Función 6 -  CallBack para enumerar los procesos y notificar
Private Function CallBackEnumProcessNotify( _
    ByVal handle As Long, _
    ByVal lParam As Long) As Boolean
    
    Dim lCountP As Long
    
    lCountP = Me.CountProcess
    
    If mCountLP <> lCountP Then
       mCountLP = lCountP
       RaiseEvent ChangeProcessList
    End If
End Function


' Función 5 -  CallBack para enumerar las ventanas y notificar
Private Function CallBackEnumWindowNotify( _
    ByVal handle As Long, _
    ByVal lParam As Long) As Boolean
    
    Dim strCaption As String
    
    strCaption = Trim(GetWindowsTaskBar(handle))
    
    If strCaption <> vbNullString Then
        mCountNW = mCountNW + Len(strCaption)
        DoEvents
    End If
    
End Function


' Función 4 - Timer para notificar cuando cambia la lista de ventanas de la barra de tareas usando la función de arriba > CallBackEnumWindowNotify
Private Function TimerProcCheckWindows( _
       ByVal hwnd As Long, _
       ByVal tMsg As Long, _
       ByVal TimerID As Long, _
       ByVal tickCount As Long) As Long
                
       mCountNW = 0
       EnumWindows scb_SetCallbackAddr(2, 5), 1
       
       If iWl <> mCountNW Then
          RaiseEvent ChangeWindowList
          iWl = mCountNW
          mCountNW = 0
       End If
End Function

' Función 3 - CallBack para enumerar las ventanas y cerrarlas a partir del nombre de clase
Private Function mCallBackCloseByClassName( _
    ByVal handle As Long, _
    ByVal lParam As Long) As Boolean
    
    Dim sClassName As String
    
    sClassName = GetClassWindowName(handle)
    
    If Trim(LCase(mClassName)) = Trim(LCase(sClassName)) Then
       ' lanzar el evento para ver las ventanas que se cerraron
       RaiseEvent CloseWindows(Me.GetWindowCaption(handle), handle, mClassName)
       Me.CloseWindowByHandle handle ' cerrar
    End If

End Function

' Función 2 -  CallBack que enumera las ventanas Child
Private Function CallBackEnumWindowChild( _
    ByVal handle As Long, _
    ByVal lParam As Long) As Boolean

    mEnumWindowChild handle
    
    CallBackEnumWindowChild = True
End Function

' Función 1 -  CallBack para enumerar las ventanas de la barra de tareas
Private Function CallBackEnumWindow( _
    ByVal handle As Long, _
    ByVal lParam As Long) As Boolean
    
    mEnumWindow handle
    
    CallBackEnumWindow = True
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Fin de Funciones CallBack :
' importante : No colocar código en este bloque, ni a continuación , hacerlo encima del la primera función del comienzo del bloque

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
