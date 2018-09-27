VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USER ADMINISTRATION MODULE"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7890
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Users"
      ForeColor       =   &H8000000E&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000D&
         Caption         =   "User Details"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   180
         TabIndex        =   13
         Top             =   3480
         Width           =   4455
         Begin VB.ComboBox cmbRight 
            DataField       =   "uRight"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "frmUsers.frx":000C
            Left            =   1740
            List            =   "frmUsers.frx":001C
            TabIndex        =   9
            Top             =   2100
            Width           =   2175
         End
         Begin VB.TextBox txtUsername 
            DataField       =   "userN"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            MaxLength       =   49
            TabIndex        =   7
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtID 
            DataField       =   "ID"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   49
            TabIndex        =   6
            Top             =   660
            Width           =   915
         End
         Begin VB.TextBox txtPassword 
            DataField       =   "passW"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            MaxLength       =   49
            TabIndex        =   8
            Top             =   1620
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Username:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   420
            TabIndex        =   18
            Top             =   1200
            Width           =   1155
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   420
            TabIndex        =   16
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Rights:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   720
            TabIndex        =   15
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Password:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Controls"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   4800
         TabIndex        =   11
         Top             =   3480
         Width           =   2655
         Begin VB.CommandButton btnUpdate 
            Caption         =   "UPDATE"
            Height          =   435
            Left            =   600
            TabIndex        =   4
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "DELETE"
            Height          =   435
            Left            =   600
            TabIndex        =   5
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton btnAdd 
            Caption         =   "NEW"
            Height          =   435
            Left            =   600
            TabIndex        =   3
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame fraCLSearch 
         BackColor       =   &H8000000D&
         Caption         =   "Users"
         ForeColor       =   &H8000000E&
         Height          =   3075
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   7275
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   3900
            TabIndex        =   2
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cmbSearch 
            Height          =   315
            ItemData        =   "frmUsers.frx":0054
            Left            =   1920
            List            =   "frmUsers.frx":0064
            TabIndex        =   1
            Text            =   "Username"
            Top             =   600
            Width           =   1695
         End
         Begin VB.Data Data1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            Connect         =   "Access"
            DatabaseName    =   "C:\Program Files\GENESIS\dbUsers.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   3180
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblUsers"
            Top             =   2580
            Width           =   1155
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmUsers.frx":0088
            Height          =   1275
            Left            =   240
            OleObjectBlob   =   "frmUsers.frx":009C
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1080
            Width           =   6795
         End
         Begin VB.Label lblSearch2 
            BackColor       =   &H8000000D&
            Caption         =   "Search by:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   2280
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblSearch1 
            BackColor       =   &H8000000D&
            Caption         =   "Search for:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4380
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblClientID 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1020
      TabIndex        =   17
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()

'-Deletes selected User record-

On Error GoTo mistake
X = MsgBox("The selected record will be permanently deleted. Are you sure?", vbOKCancel, "Warning!")
If X = vbOK Then
Data1.Recordset.Delete
Data1.Refresh
End If
mistake:

'--------------------------------

End Sub

Private Sub btnUpdate_Click()

'-Updates selected User record-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------------

End Sub


Private Sub btnAdd_Click()

'-Adds a new User record-

On Error GoTo mistake
Data1.Recordset.AddNew
txtUsername.SetFocus
Exit Sub
mistake:
MsgBox "An error occured during data loading. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'------------------------

End Sub



Private Sub Form_Load()

'-Loads Data to the form controls-----

On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "/dbUsers.mdb"
Data1.Refresh
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'-------------------------------------

End Sub

Private Sub Form_Unload(Cancel As Integer)

'-Updates User records-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------

End Sub

Private Sub txtName_Change()

'-Performs a search for the specified User-

On Error GoTo mistake:
If txtName.Text <> "" Then
    Select Case cmbSearch.Text
        Case "ID"
            Data1.RecordSource = "select * from tblUsers where ID = " & txtName.Text & ";"
            Data1.Refresh
        Case "Username"
            Data1.RecordSource = "select * from tblUsers where userN like '" & txtName.Text & "*" & "';"
            Data1.Refresh
        Case "Password"
            Data1.RecordSource = "select * from tblUsers where passW like '" & txtName.Text & "*" & "';"
            Data1.Refresh
        Case "Rights"
            Data1.RecordSource = "select * from tblUsers where uRight like '" & txtName.Text & "*" & "';"
            Data1.Refresh
    End Select
Else
    Data1.RecordSource = "select * from tblUsers;"
    Data1.Refresh
End If
Exit Sub

mistake:
MsgBox "Invalid input. Please insert a correct value.", vbExclamation, "Caution!"

'----------------------------------------------------------------------------------------


End Sub
