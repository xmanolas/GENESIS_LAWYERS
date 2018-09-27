VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientSearch 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENT SEARCH MODULE"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "frmClientSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   4920
      TabIndex        =   3
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Frame fraCLSearch 
      BackColor       =   &H8000000D&
      Caption         =   "Client Search"
      ForeColor       =   &H8000000E&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmClientSearch.frx":000C
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "frmClientSearch.frx":0020
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   10215
      End
      Begin VB.Data Data1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\GENESIS\dbGenesis.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   4860
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblClients"
         Top             =   3960
         Width           =   1155
      End
      Begin VB.ComboBox cmbSearch 
         Height          =   315
         ItemData        =   "frmClientSearch.frx":09F3
         Left            =   3540
         List            =   "frmClientSearch.frx":0A12
         TabIndex        =   1
         Text            =   "Surname"
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   5520
         TabIndex        =   2
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label lblSearch1 
         BackColor       =   &H8000000D&
         Caption         =   "Search for:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblSearch2 
         BackColor       =   &H8000000D&
         Caption         =   "Search by:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3900
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.TextBox txtClientID 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtSurname 
      DataField       =   "Surname"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtFirstname 
      DataField       =   "Firstname"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6900
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmClientSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()

'-Sends Client data to Main module-

mdlMain.ClientID = txtClientID.Text
mdlMain.ClientFirstname = txtFirstname.Text
mdlMain.ClientSurname = txtSurname.Text
Unload Me

'--------------------------------

End Sub

Private Sub Form_Load()

'-Loads Data to the form controls-----

On Error GoTo mistake
mdlMain.ClientID = ""
Data1.DatabaseName = mdlMain.ApplicPath & "/dbGenesis.mdb"
Data1.Refresh
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'-------------------------------------

End Sub


Private Sub txtName_Change()

'-Sends SQL commands to database. It works as a search engine for client records-

On Error GoTo mistake:
If txtName.Text <> "" Then
    Select Case cmbSearch.Text
        Case "ID"
            Data1.RecordSource = "select * from tblClients where ID = " & txtName.Text & ";"
            Data1.Refresh
        Case Else
            Data1.RecordSource = "select * from tblClients where " & cmbSearch.Text & " like '" & txtName.Text & "*" & "';"
            Data1.Refresh
    End Select
Else
    Data1.RecordSource = "select * from tblClients;"
    Data1.Refresh
End If
Exit Sub

mistake:
MsgBox "Invalid input. Please insert a correct value.", vbExclamation, "Caution!"

'----------------------------------------------------------------------------------------

End Sub
