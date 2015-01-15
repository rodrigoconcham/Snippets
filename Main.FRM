VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   Caption         =   "snippets"
   ClientHeight    =   7185
   ClientLeft      =   1575
   ClientTop       =   1830
   ClientWidth     =   9915
   DrawMode        =   1  'Blackness
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   HelpContextID   =   20
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7185
   ScaleWidth      =   9915
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   6825
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Description     =   "Exit"
            Object.ToolTipText     =   "Exit Code bank"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete"
            Object.ToolTipText     =   "Delete routine"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New routine"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit routine"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy to clipboard"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMain 
      Left            =   3300
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0442
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":055E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":067A
            Key             =   "code"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0796
            Key             =   "codered"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08B2
            Key             =   "note"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":09CE
            Key             =   "vb"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0AEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0C02
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0D1A
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0E32
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F4A
            Key             =   "repair"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1062
            Key             =   "help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":117A
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1292
            Key             =   "save"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":13AA
            Key             =   "abort"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabCode 
      Height          =   5775
      Left            =   3360
      TabIndex        =   11
      Top             =   480
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   10186
      TabFixedWidth   =   2646
      TabFixedHeight  =   706
      ImageList       =   "ilsMain"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Descripcion"
            Object.ToolTipText     =   "Functional description"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Codigo"
            Object.ToolTipText     =   "VB source code"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRoute 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtCodeWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   0
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtCodeWindow 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Index           =   1
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox txtCodeWindow 
      Height          =   5295
      Index           =   2
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox txtCodeWindow 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Index           =   3
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   840
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   60
      ScaleHeight     =   6135
      ScaleWidth      =   3135
      TabIndex        =   1
      Top             =   540
      Width           =   3135
      Begin MSComctlLib.TreeView tvwMain 
         Height          =   5835
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   10292
         _Version        =   393217
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ilsMain"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog dlgFileOpenClose 
      Left            =   3900
      Top             =   195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCategory 
      Height          =   195
      Index           =   1
      Left            =   4500
      TabIndex        =   10
      Top             =   420
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblCategory 
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lblCounter 
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Routine"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   9480
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   9480
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Form level variables
 Dim fvDisplayOnly As Boolean
 Dim fvEdit As Boolean

 Dim dbName As Database
 Dim rsCategories As Recordset
 Dim rsRoutines As Recordset
 Dim rsFiltered As Recordset
 Dim lActivarForm As Boolean
 
 Dim CurrentNode As Node

Public Sub ExitProgram()
'Close CodeBank Tables
  Call CloseTables

'Loop thru the forms collection and unload all forms from memory
  Dim frm As Form
  For Each frm In Forms
     Unload frm          ' deactivate the form
     Set frm = Nothing   ' remove the object from memory
  Next

End Sub

Public Sub AddNode(strCategoryName As String, strRoutineKey As String, strRoutineName As String)
' Adds a child node to the treeview
' Input: Category name, Routine key, Routine name
    
  On Error Resume Next
    
  Dim ndNewNode As Node
  Set ndNewNode = tvwMain.Nodes.Add(strCategoryName, tvwChild, strRoutineKey, strRoutineName, 3, 4)

End Sub

Public Sub CloseTables()
  On Error Resume Next
    rsRoutines.Close
  On Error Resume Next
    rsCategories.Close
  On Error Resume Next
    rsFiltered.Close
  dbName.Close

End Sub

Public Sub CopyToClipboard()
'Exit if nothing in code window
  If Len(txtCodeWindow(1).Text) < 1 Then
     MsgBox "Please select a routine to copy.", vbExclamation + vbOKOnly, "Nothing to copy"
     Exit Sub
  End If
  
'Clear Clipboard.
  Clipboard.Clear

'Select Text in txtCodeWindow(1) & copy to clipboard
  Clipboard.SetText txtCodeWindow(1).Text, vbCFText
  
'Display MsgBox
  MsgBox "The selected routine is now on the Clipboard.", vbInformation + vbOKOnly, "Code copied"

End Sub

Public Sub DeleteRoutine()
'Dim error string
  Dim strErr As String

'Set error trap in case no current node
  On Error GoTo NodeError
  
'If Category selected then exit
  If Left(CurrentNode.Key, 8) = "Category" Then
     MsgBox "Please select a routine to delete.", vbExclamation + vbOKOnly, "Nothing to delete"
     Exit Sub
  End If
       
'If user clicks the No button, Exit sub
  If MsgBox("Do you really want to delete the " & CurrentNode.Text & " routine?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm delete") = vbNo Then Exit Sub
  
'SQL to point to record
  Dim SelectString As String
  SelectString = "Select * from Routine where [Counter]= " + lblCounter
  Set rsRoutines = dbName.OpenRecordset(SelectString)

'Delete record
  rsRoutines.Delete
     
'Save parent node
  Dim DeletedNodeParent As String
  DeletedNodeParent = tvwMain.SelectedItem.Parent.Key

'Remove node
  tvwMain.Nodes.Remove tvwMain.SelectedItem.Key

'Select parent node
  tvwMain.Nodes(DeletedNodeParent).Selected = True
  
'Clear text boxes
  txtCodeWindow(0).Text = ""
  txtCodeWindow(1).Text = ""

'Exit before error handler
  Exit Sub

NodeError:
  MsgBox "Please select a routine to delete.", vbExclamation + vbOKOnly, "Nothing to delete"
  Exit Sub

End Sub

Public Sub EditRoutine()
'Dim error string
  Dim strErr As String

'Set error trap in case no current node
  On Error GoTo NodeError
  
'If Category selected then exit
  If Left(CurrentNode.Key, 8) = "Category" Then
     MsgBox "Please select a routine to edit.", vbExclamation + vbOKOnly, "Nothing to edit"
     Exit Sub
  End If
     
'Reset error trap
  On Error Resume Next

'Copy text to be edited
  txtRoute.Text = CurrentNode.Text
  txtCodeWindow(2).Text = txtCodeWindow(0).Text
  txtCodeWindow(3).Text = txtCodeWindow(1).Text

'Set labels
  lblCategory(0) = CurrentNode.Parent.Text
  lblCategory(1) = CurrentNode.Parent.Key

'Hide treeview
  Picture1.Visible = False
  
'Hide description & code snippet display windows
  txtCodeWindow(0).Visible = False
  txtCodeWindow(1).Visible = False
   
'Show Routine Input Box
  txtRoute.Visible = True
  txtRoute.SetFocus
  
'Show Save/Abort command buttons
  Toolbar1.Buttons("Save").Visible = True
  Toolbar1.Buttons("Stop").Visible = True
   
'Show description & code snippet display windows
  txtCodeWindow(2).Visible = True
  txtCodeWindow(3).Visible = True
   
'This code puts the selected tab's text container on top
  txtCodeWindow(tabCode.SelectedItem.Index + 1).ZOrder 0
   
'Set Editing Flags
  fvEdit = True
  fvDisplayOnly = False
  
'Disable routine buttons
  Toolbar1.Buttons("Delete").Enabled = False
  Toolbar1.Buttons("Edit").Enabled = False
  Toolbar1.Buttons("Copy").Enabled = False
  Toolbar1.Buttons("Repair").Enabled = False
  
'Exit before error
  Exit Sub
  
NodeError:
   MsgBox "Please select a routine to edit.", vbExclamation + vbOKOnly, "Nothing to edit"

End Sub

Public Sub NewRoutine()
'Error trap if no node currently selected
  On Error GoTo NodeError
  
'If Child selected then select Parent
  If Left(CurrentNode.Key, 7) = "Routine" Then
     lblCategory(0) = CurrentNode.Parent.Text
     lblCategory(1) = CurrentNode.Parent.Key
  Else
     lblCategory(0) = CurrentNode.Text
     lblCategory(1) = CurrentNode.Key
  End If
     
'Reset error trap
  On Error Resume Next

'Hide treeview
  Picture1.Visible = False
  
'Hide description & code snippet display windows
  txtCodeWindow(0).Visible = False
  txtCodeWindow(1).Visible = False
   
'Show Routine Input Box
  txtRoute.Visible = True
  txtRoute.SetFocus
  
'Show Save/Abort command buttons
  Toolbar1.Buttons("Save").Visible = True
  Toolbar1.Buttons("Stop").Visible = True
   
'Show description & code snippet display windows
   txtCodeWindow(2).Visible = True
   txtCodeWindow(3).Visible = True
   
'This code puts the selected tab's text container on top
   txtCodeWindow(tabCode.SelectedItem.Index + 1).ZOrder 0
   
'Set flags to new
   fvEdit = False
   fvDisplayOnly = False
   
'Disable routine buttons
  Toolbar1.Buttons("Delete").Enabled = False
  Toolbar1.Buttons("Edit").Enabled = False
  Toolbar1.Buttons("Copy").Enabled = False
  Toolbar1.Buttons("Repair").Enabled = False
   
'Exit before error handler
  Exit Sub
   
NodeError:
   MsgBox "Please select a category.", vbExclamation + vbOKOnly, "Unable to add routine"
  
End Sub



Public Sub Abort()
'Show routine list box
  Picture1.Visible = True
   
'Show description & code snippet display windows
  txtCodeWindow(0).Visible = True
  txtCodeWindow(1).Visible = True
   
'Clear info in code display
  If fvEdit = True Then
    txtCodeWindow(0).Text = ""
    txtCodeWindow(1).Text = ""
  End If

'Hide Routine Input Box
  txtRoute.Visible = False
  
'Hide Save/Abort command buttons
  Toolbar1.Buttons("Save").Visible = False
  Toolbar1.Buttons("Stop").Visible = False

'Hide description & code snippet input
  txtCodeWindow(2).Visible = False
  txtCodeWindow(3).Visible = False
   
'Clear input boxes
  txtRoute.Text = ""
  txtCodeWindow(2).Text = ""
  txtCodeWindow(3).Text = ""
  
'Set tab to 0
  txtCodeWindow(tabCode.SelectedItem.Index - 1).ZOrder 0

'Reset flags to display
  fvDisplayOnly = True
  fvEdit = False

'Enable routine buttons
  Toolbar1.Buttons("Delete").Enabled = True
  Toolbar1.Buttons("Edit").Enabled = True
  Toolbar1.Buttons("Copy").Enabled = True
  
'SetFocus to treeview
  tvwMain.SetFocus
  
'Select parent node
  tvwMain.Nodes(lblCategory(1).Caption).Selected = True

End Sub

Public Sub Save()
'Exit if any data missing
  Dim strErr As String
  
  If Len(txtRoute.Text) < 1 Then
     MsgBox "Please enter routine name.", vbExclamation + vbOKOnly, "Validation error"
     txtRoute.SetFocus
     Exit Sub
  End If
  
  If Len(txtCodeWindow(2).Text) < 1 Then
     MsgBox "Please enter description.", vbExclamation + vbOKOnly, "Validation error"
     txtCodeWindow(2).SetFocus
     Exit Sub
  End If
  
  If Len(txtCodeWindow(3).Text) < 1 Then
     MsgBox "Please enter code snippet.", vbExclamation + vbOKOnly, "Validation error"
     txtCodeWindow(3).SetFocus
     Exit Sub
  End If

'Enable error trapping
  On Error GoTo AddError
   
'Set Edit Flag
  If fvEdit = True Then
    'SQL to find routine
     Dim SelectString As String
     SelectString = "Select * from Routine where [Counter]= " + lblCounter
     Set rsFiltered = dbName.OpenRecordset(SelectString)
    'Edit record
     rsFiltered.Edit
     rsFiltered.Fields("Routine") = txtRoute.Text
     rsFiltered.Fields("Description") = txtCodeWindow(2).Text
     rsFiltered.Fields("Snippet") = txtCodeWindow(3).Text
     rsFiltered.Update
     
   'Display MsgBox
     MsgBox "The " + txtRoute.Text + " code snippet has been saved.", vbInformation + vbOKOnly, "Code saved"
     
   'Change Node text
      tvwMain.Nodes(CurrentNode.Key).Text = txtRoute.Text
   
   'Return to View Model
     Call Abort
  
  Else
    'Add new record
     rsRoutines.AddNew
     rsRoutines.Fields("Category") = lblCategory(0)
     rsRoutines.Fields("Routine") = txtRoute.Text
     rsRoutines.Fields("Description") = txtCodeWindow(2).Text
     rsRoutines.Fields("Snippet") = txtCodeWindow(3).Text
     rsRoutines.Update
    'Move to new record
     rsRoutines.Bookmark = rsRoutines.LastModified
  
   'Display MsgBox
     MsgBox "The " + txtRoute.Text + " code snippet has been added to VB Code Bank.", vbInformation + vbOKOnly, "Code saved"

   'Clear input boxes
     txtRoute.Text = ""
     txtCodeWindow(2).Text = ""
     txtCodeWindow(3).Text = ""
  
   'Add new routine to treeview
     Call AddNode(lblCategory(1), "Routine" & rsRoutines("counter"), rsRoutines("Routine"))
  End If
  
'Disable error trapping & exit
  On Error GoTo 0
  Exit Sub

AddError:
 'Get error number
  strErr = "VB error " & Err.Number & vbCrLf & Err.Description
  MsgBox "Unable to add/edit routine." + vbCrLf + vbCrLf + strErr, vbExclamation + vbOKOnly, "Database error"
 
End Sub

Private Sub Form_Load()
'*** Code added by HelpWriter ***
   SetAppHelp Me.hWnd
    lActivarForm = False ' para verificar por primera vez la ruta de base

'Center on the screen
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2

' This code puts the selected tab's text container on top.
  txtCodeWindow(tabCode.SelectedItem.Index - 1).ZOrder 0

'Set form variable to switch between display & insert new
  fvDisplayOnly = True

'Set Edit Flag
  fvEdit = False
  
'Open Category Table with code
  Call OpenTables
                                                                  
'Load tree view
  Call LoadTreeView
  
'Disable Routine buttons until a routine is selected
  Toolbar1.Buttons("Delete").Enabled = False
  Toolbar1.Buttons("Edit").Enabled = False
  Toolbar1.Buttons("Copy").Enabled = False
  Toolbar1.Buttons("New").Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then      'User chose close from control menu
   'If user clicks the No button, stop Exit
     If MsgBox("Quieres salir de Snippets?", vbQuestion + vbYesNo + vbDefaultButton2, "Salir de Snippets") = vbNo Then
        Cancel = True
        Exit Sub
     End If
    
   'Call cleanup code
     Call ExitProgram
     
  End If


End Sub

Private Sub mnuCopy_Click()
  Call CopyToClipboard
  
End Sub

Private Sub mnuDeleteRoutine_Click()
  Call DeleteRoutine
  
End Sub

Private Sub mnuEdit_Click()
' Before the menu is dropped down this routine checks that there is a
' routine selected, if yes, then enable the menu items

  'mnuEditRoutine.Enabled = False
  'mnuDeleteRoutine.Enabled = False
  'mnuCopy.Enabled = False
  
'If tree view has not been selected then an error will occur
  On Error GoTo DisplayOnly
  
  If Left(CurrentNode.Key, 7) = "Routine" Then
      'mnuEditRoutine.Enabled = True
      'mnuDeleteRoutine.Enabled = True
      'mnuCopy.Enabled = True
  End If

DisplayOnly:
' Before the menu is dropped down this routine checks if in
' Edit or Display mode, if Editing then disable the menu items
  If fvDisplayOnly = False Then
     'mnuEditRoutine.Enabled = False
     'mnuDeleteRoutine.Enabled = False
     'mnuCopy.Enabled = False
  End If

End Sub

Private Sub mnuEditRoutine_Click()
  Call EditRoutine
  
End Sub

Private Sub mnuFile_Click()
' Before the menu is dropped down this routine checks if in
' Edit mode


End Sub

Private Sub mnuFileExit_Click()
 'If user clicks the No button, stop Exit
 If MsgBox("Quieres salir de Snippets?", vbQuestion + vbYesNo + vbDefaultButton2, "Salir de Snippets") = vbNo Then
  Call ExitProgram
 End If
End Sub

Private Sub mnuHelpContents_Click()
   Call ShowHelpContents

End Sub

Private Sub mnuInsert_Click()
' Before the menu is dropped down this routine checks if in
' Edit mode

  'If fvDisplayOnly = False Then
    '  mnuNewRoutine.Enabled = False
   '   mnuNewCategory.Enabled = False
  'Else
   '   mnuNewRoutine.Enabled = True
    '  mnuNewCategory.Enabled = True
  'End If

End Sub

Private Sub mnuNewCategory_Click()
'Input box to ask for category name
  Dim CategoryName As String
  Do
    CategoryName = InputBox("Enter category name", "New category")
    If CategoryName = "" Then Exit Sub     'Cancel button pressed
  Loop Until Len(CategoryName) > 0
   
'Enable error trapping
  On Error GoTo AddError
   
'Add new record
  rsCategories.AddNew
  rsCategories.Fields("Category") = CategoryName
  rsCategories.Update

'Move to new record
  rsCategories.Bookmark = rsCategories.LastModified
  
'Add new category to treeview
  tvwMain.Nodes.Add , , "Category" & rsCategories("counter"), rsCategories("Category"), "closed", "open"

'Disable error trapping
  On Error GoTo 0
  Exit Sub

AddError:
 'Get error number
  Dim strErr As String
  strErr = "VB error " & Err.Number & vbCrLf & Err.Description
  MsgBox "Unable to add category." + vbCrLf + vbCrLf + strErr, vbExclamation + vbOKOnly, "Database error"
     
End Sub

Private Sub mnuHelpAbout_Click()
'Load the About box

End Sub

Private Sub mnuNewRoutine_Click()
  Call NewRoutine

End Sub

Private Sub mnuRepairdb_Click()
   

End Sub

Private Sub tabCode_Click()
' This code puts the selected tab's text container on top
  If fvDisplayOnly = True Then
     txtCodeWindow(tabCode.SelectedItem.Index - 1).ZOrder 0
  Else
     txtCodeWindow(tabCode.SelectedItem.Index + 1).ZOrder 0
  End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
'Handle button clicks
  Select Case Button.Key
    Case Is = "Exit"
     'If user clicks the No button, stop Exit
       If MsgBox("Quieres salir de Snippets?", vbQuestion + vbYesNo + vbDefaultButton2, "Salir de Snippets") = vbYes Then
        Call ExitProgram
      End If
       
    Case Is = "Delete"
       Call DeleteRoutine
    Case Is = "Edit"
       Call EditRoutine
    Case Is = "New"
       Call NewRoutine
    Case Is = "Copy"
       Call CopyToClipboard
    Case Is = "Help"
       Call ShowHelpContents
    Case Is = "Save"
       Call Save
    Case Is = "Stop"
       Call Abort

  End Select
  
End Sub

Private Sub tvwMain_BeforeLabelEdit(Cancel As Integer)
'This prevents editing nodes in the tree view control
  Cancel = True
  
End Sub

Private Sub tvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
 'Clear code display windows
  txtCodeWindow(0) = ""
  txtCodeWindow(1) = ""

End Sub

Private Sub tvwMain_KeyUp(KeyCode As Integer, Shift As Integer)
'Exit if nothing selected in tree view
  If tvwMain.SelectedItem Is Nothing Then Exit Sub
                                 
'Enable routine buttons
  Toolbar1.Buttons("Delete").Enabled = True
  Toolbar1.Buttons("Edit").Enabled = True
  Toolbar1.Buttons("Copy").Enabled = True

'Enable error trapping
  On Error GoTo CodeError
   
'SQL to Find Description & Snippet for the displayed Routine
  Dim SelectString As String        'Set SQL query
  SelectString = "Select * from Routine where [Routine]= '" + tvwMain.SelectedItem + "'"
  
  Dim rsName As Recordset ' record set object
  Set rsName = dbName.OpenRecordset(SelectString)
     
'Load Description & Snippet boxes
  lblCounter = rsName.Fields("Counter").Value
  txtCodeWindow(0) = rsName.Fields("Description").Value
  txtCodeWindow(1) = rsName.Fields("Snippet").Value
                                 
'Disable error trapping
  On Error GoTo 0
  rsName.Close
  Exit Sub

CodeError:
 'Clear code windows
  txtCodeWindow(0) = ""
  txtCodeWindow(1) = ""

'Disable routine buttons
  Toolbar1.Buttons("Delete").Enabled = False
  Toolbar1.Buttons("Edit").Enabled = False
  Toolbar1.Buttons("Copy").Enabled = False

End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Exit if nothing selected in tree view
  If tvwMain.SelectedItem Is Nothing Then Exit Sub
                                 
'Enable routine buttons
  Toolbar1.Buttons("Delete").Enabled = True
  Toolbar1.Buttons("Edit").Enabled = True
  Toolbar1.Buttons("Copy").Enabled = True

'Enable error trapping
  On Error GoTo CodeError
   
'SQL to Find Description & Snippet for the displayed Routine
  Dim SelectString As String        'Set SQL query
  SelectString = "Select * from Routine where [Routine]= '" + tvwMain.SelectedItem + "'"
  
  Dim rsName As Recordset ' record set object
  Set rsName = dbName.OpenRecordset(SelectString)
     
'Load Description & Snippet boxes
  lblCounter = rsName.Fields("Counter").Value
  txtCodeWindow(0) = rsName.Fields("Description").Value
  txtCodeWindow(1) = rsName.Fields("Snippet").Value
                                 
'Disable error trapping
  On Error GoTo 0
  rsName.Close
  Exit Sub

CodeError:
 'Clear code windows
  txtCodeWindow(0) = ""
  txtCodeWindow(1) = ""
'Enable routine buttons
  Toolbar1.Buttons("Delete").Enabled = False
  Toolbar1.Buttons("Edit").Enabled = False
  Toolbar1.Buttons("Copy").Enabled = False
 
End Sub

Public Sub OpenTables()


Dim i As Integer


  ' Set dbName = OpenDatabase(App.Path + "\codebank.mdb")
  
  
     If Not lActivarForm Then
        lActivarForm = True
        sRutaIni = App.Path & "\" & "Snippets.ini"
        ' validar ruta de base
        sRutaBase = LeerIni("PARAMETROS", "RUTABASE")
        If sRutaBase = "" Then ' si no existe
            FrmEspera.Show 1
        End If
        
 'obtener ruta de base


For i = Len(sRutaBase) To 1 Step -1
   If Mid(sRutaBase, i, 1) = "\" Then
        sruta = Left(sRutaBase, i)
        Exit For
   End If
Next
  
End If
  
  'Set dbName = OpenDatabase(sRutaBase + "\codebank.mdb")
  
  Set gwsWorkspace = DBEngine.CreateWorkspace("NewJetWorkspace", "Admin", "", dbUseJet)
  Set dbName = gwsWorkspace.OpenDatabase(sRutaBase, False, False)
  
  
  Set rsCategories = dbName.OpenRecordset("SELECT * FROM Category " & "ORDER BY Category") ', dbOpenSnapshot)
  Set rsRoutines = dbName.OpenRecordset("SELECT * FROM Routine " & "ORDER BY Category, Routine") ', dbOpenSnapshot)

End Sub

Public Sub LoadTreeView()
'Load categories into tree view
  Do While Not rsCategories.EOF
  'Add category nodes
     tvwMain.Nodes.Add , , "Category" & rsCategories("counter"), rsCategories("Category"), "closed", "open"
     rsCategories.MoveNext
  Loop
   
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As Node)
  Toolbar1.Buttons("New").Enabled = True
    
  Set CurrentNode = Node
  
 'Exit If Child selected
  If Left(CurrentNode.Key, 7) = "Routine" Then
     Exit Sub
  Else
     lblCategory(0) = CurrentNode.Text
     lblCategory(1) = CurrentNode.Key
  End If
  
 'Add routine nodes
  If Len(lblCategory(0)) > 0 Then
    'Filter the routine table for current category
     rsRoutines.Filter = "Category = '" & lblCategory(0) & "'"
     Set rsFiltered = rsRoutines.OpenRecordset

    'Loop through filtered recordset
     Do While Not rsFiltered.EOF
        'Add nodes for routines
         Call AddNode(lblCategory(1), "Routine" & rsFiltered("counter"), rsFiltered("Routine"))
         rsFiltered.MoveNext
     Loop
              
    'Expand current node
     CurrentNode.Expanded = True
     
  End If
  
End Sub


