VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocuments 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASE DOCUMENTS"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   Icon            =   "frmDocuments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      DataField       =   "RelatedDocument"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCaseID 
      DataField       =   "CaseID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   5820
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Case Documents"
      ForeColor       =   &H8000000E&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9075
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000D&
         Caption         =   "Selected Document"
         ForeColor       =   &H8000000E&
         Height          =   1515
         Left            =   120
         TabIndex        =   13
         Top             =   2940
         Width           =   6135
         Begin VB.CommandButton btnOpen 
            Caption         =   "OPEN DOCUMENT"
            Height          =   555
            Left            =   2340
            TabIndex        =   14
            Top             =   780
            Width           =   1515
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "None"
            DataField       =   "RelatedDocument"
            DataSource      =   "Data1"
            ForeColor       =   &H8000000E&
            Height          =   435
            Left            =   60
            TabIndex        =   15
            Top             =   300
            Width           =   6015
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Controls"
         ForeColor       =   &H8000000E&
         Height          =   3855
         Left            =   6360
         TabIndex        =   9
         Top             =   600
         Width           =   2595
         Begin VB.CommandButton Command2 
            Caption         =   "UPDATE"
            Height          =   435
            Left            =   540
            TabIndex        =   3
            Top             =   1620
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "REMOVE"
            Height          =   435
            Left            =   540
            TabIndex        =   4
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton btnAdd 
            Caption         =   "ADD"
            Height          =   435
            Left            =   540
            TabIndex        =   2
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame fraCLSearch 
         BackColor       =   &H8000000D&
         Caption         =   "Document Search"
         ForeColor       =   &H8000000E&
         Height          =   2295
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   6135
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmDocuments.frx":000C
            Height          =   915
            Left            =   180
            OleObjectBlob   =   "frmDocuments.frx":0020
            TabIndex        =   16
            Top             =   780
            Width           =   5835
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
            Left            =   2580
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblCaseDocuments"
            Top             =   1800
            Width           =   1155
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   3480
            TabIndex        =   1
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label lblSearch1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Search for document by name (or part of the document's path)"
            ForeColor       =   &H8000000E&
            Height          =   435
            Left            =   1020
            TabIndex        =   8
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Label lblCaseID 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4500
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "CASE ID:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      Class           =   "Package"
      DataSource      =   "Data1"
      DisplayType     =   1  'Icon
      Height          =   1515
      Left            =   2580
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "frmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()

'-Opens the dialog control and accepts a document's path from the user-

Data1.Recordset.AddNew
CommonDialog1.ShowOpen
txtPath.Text = CommonDialog1.FileName
txtCaseID.Text = lblCaseID.Caption

'----------------------------------------------------------------------

End Sub


Private Sub btnOpen_Click()

'-Opens Related Document----

On Error GoTo mistake
OLE1.CreateLink (txtPath.Text)
OLE1.DoVerb (1)
Exit Sub
mistake:
MsgBox "Invalid file type!", vbCritical, "Error!"

'---------------------------

End Sub

Private Sub Command1_Click()

'-Deletes selected Document record-

On Error GoTo mistake
X = MsgBox("The selected record will be permanently deleted. Are you sure?", vbOKCancel, "Warning!")
If X = vbOK Then
Data1.Recordset.Delete
Data1.Refresh
End If
mistake:

'--------------------------------

End Sub

Private Sub Command2_Click()

'-Updates selected Case Document record-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------------------

End Sub


Private Sub Form_Load()

'-Tries to read and load CaseID value from related variable in mdlMain-

On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "/dbGenesis.mdb"
If mdlMain.CaseID <> "" Then
lblCaseID.Caption = mdlMain.CaseID
Data1.RecordSource = "select * from tblCaseDocuments where CaseID like '" & lblCaseID.Caption & "';"
Data1.Refresh
End If
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'----------------------------------------------------------------------

End Sub

Private Sub Form_Unload(Cancel As Integer)

'-Updates selected Case Document record-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"


'--------------------------------------

End Sub

Private Sub txtName_Change()

'-Sends SQL commands to database. It works as a search engine for Case Documents records-

On Error GoTo mistake:
If txtName.Text <> "" Then
    Data1.RecordSource = "select * from tblCaseDocuments where RelatedDocument like '" & "*" & txtName.Text & "*" & "'and CaseID like '" & lblCaseID.Caption & "' ;"
    Data1.Refresh
Else
    Data1.RecordSource = "select * from tblCaseDocuments where CaseID like '" & lblCaseID.Caption & "';"
    Data1.Refresh
End If
Exit Sub
mistake:
MsgBox "Invalid input. Please insert a correct value.", vbExclamation, "Caution!"

'-----------------------------------------------------------------------------------------

End Sub
