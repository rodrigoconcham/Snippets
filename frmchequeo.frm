VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmchequeo 
   BackColor       =   &H00D5B29C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TERCOP"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvc 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hwnd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ClasName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEMA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   1785
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D5B29C&
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   1575
      Left            =   480
      Top             =   180
      Width           =   3015
   End
End
Attribute VB_Name = "frmchequeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmchequeo
' DateTime  : 17/10/2014 12:00
' Author    : Jbiott
' Purpose   : no ejecutar mas instancias que las deseadas de una misma aplicacion
'               el numero de instancias permitidas son las definidas en la constante
'               NroInstancias
'---------------------------------------------------------------------------------------
' referencias: mscomctl.OCX
Option Explicit
Private Const NroInstancias = 1

Private WithEvents cEnumProc As clsEnum
Attribute cEnumProc.VB_VarHelpID = -1


Private Sub cEnumProc_Error(ByVal sError As String)
    Debug.Print sError
End Sub

Private Sub cEnumProc_GetProcess( _
    ByVal sNameProcess As String, _
    ByVal SpathProcess As String, _
    ByVal HandleProcess As Long)
    
    Dim Item As ListItem
    
    Set Item = lvc.ListItems.Add(, , sNameProcess)
    Item.SubItems(1) = SpathProcess
    Item.SubItems(2) = HandleProcess
End Sub
Private Sub cEnumProc_StartEnumProcess()
    lvc.ListItems.Clear
End Sub

Private Sub cargar_procesos()
    With cEnumProc
        .EnumerateProcesses
     End With
End Sub

Private Function Nro_instancias(proceso As String) As Integer
Dim i As Integer
Dim busqueda As String
Dim c As Integer

c = 0
busqueda = proceso
For i = 1 To lvc.ListItems.Count
    If UCase(busqueda) = UCase(lvc.ListItems(i).Text) Then
        c = c + 1
    End If
Next i

Nro_instancias = c
End Function

Private Sub Form_Load()


    lvc.Sorted = True
    Set cEnumProc = New clsEnum
    cEnumProc.TopMost Me.hwnd, True
    cargar_procesos
    
    If Nro_instancias(App.EXEName & ".exe") > NroInstancias Then
        MsgBox "Por favor por razones de estabilidad. No se pueden abrir más de " & NroInstancias & " Instancias", vbInformation + vbOKOnly, "Segunda Copia"
        End
    Else
    Unload Me
    '*******************************************************************
    frmMain.Show  'aca debe ir formulario principal del proyecto'
    '*********************************************************************
    End If
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set cEnumProc = Nothing
End Sub





