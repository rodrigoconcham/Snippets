VERSION 5.00
Begin VB.Form FrmEspera 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ingrese ruta"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   6645
   Begin VB.CommandButton Cmd_cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_aceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   1920
      Pattern         =   "*.mdb"
      TabIndex        =   1
      Top             =   2520
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8E9EC&
      Caption         =   "La base de Datos se Guardará en La Siguiente Ruta"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3720
   End
End
Attribute VB_Name = "FrmEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const WS_THICKFRAME = &H40000
Const GWL_STYLE = (-16)

Private Sub Cmd_aceptar_Click()

Dim sruta   As String
    sruta = Dir1.Path & IIf(Right(Dir1.Path, 1) = "\", "", "\") & File1.List(File1.ListIndex) ' Drive1.List(Drive1.ListIndex) &
    
    sRutaBase = sruta
    Call GrabarIni("PARAMETROS", "RUTABASE", sruta)
    Unload Me



End Sub

Private Sub Cmd_cancelar_Click()
  If MsgBox("¿Seguro desea cancelar?", vbQuestion + vbYesNo, "Base") = vbYes Then
        End
    End If
End Sub





Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Refresh
    
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorRuta
    Dir1.Path = Drive1.List(Drive1.ListIndex)
    Dir1.Refresh
     Exit Sub
ErrorRuta:
    MsgBox "error"
End Sub

