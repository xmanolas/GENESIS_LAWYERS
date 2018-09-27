VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCases 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASE MANAGEMENT MODULE"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "frmCases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCLUpdate 
      BackColor       =   &H8000000D&
      Caption         =   "Case Information Update"
      ForeColor       =   &H8000000E&
      Height          =   4935
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   10695
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Controls"
         ForeColor       =   &H8000000E&
         Height          =   4575
         Left            =   5340
         TabIndex        =   29
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton btnPrint 
            Caption         =   "PRINT"
            Height          =   495
            Left            =   1860
            TabIndex        =   6
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "DELETE"
            Height          =   495
            Left            =   1860
            TabIndex        =   5
            Top             =   2460
            Width           =   1455
         End
         Begin VB.CommandButton btnUpdate 
            Caption         =   "UPDATE"
            Height          =   495
            Left            =   1860
            TabIndex        =   4
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton btnAddnew 
            Caption         =   "CREATE"
            Height          =   495
            Left            =   1860
            TabIndex        =   3
            Top             =   660
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000D&
         Caption         =   "Details"
         ForeColor       =   &H8000000E&
         Height          =   4575
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton btnShowLawyer 
            Caption         =   "DETAILS"
            Height          =   315
            Left            =   2580
            TabIndex        =   11
            Top             =   1440
            Width           =   1035
         End
         Begin VB.CommandButton btnShowClient 
            Caption         =   "DETAILS"
            Height          =   315
            Left            =   2580
            TabIndex        =   9
            Top             =   1020
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Bindings        =   "frmCases.frx":000C
            DataField       =   "StartDate"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1380
            TabIndex        =   14
            Top             =   3180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20119553
            CurrentDate     =   38718
         End
         Begin VB.CommandButton btnDocuments 
            Caption         =   "RELATED DOCUMENTS"
            Height          =   555
            Left            =   3060
            TabIndex        =   16
            Top             =   3060
            Width           =   1515
         End
         Begin VB.CommandButton btnLawyerS 
            Caption         =   "SEARCH"
            Height          =   315
            Left            =   1380
            TabIndex        =   31
            Top             =   1440
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton btnClientS 
            Caption         =   "SEARCH"
            Height          =   315
            Left            =   1380
            TabIndex        =   30
            Top             =   1020
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.ComboBox cmbStatus 
            DataField       =   "Status"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "frmCases.frx":002F
            Left            =   1380
            List            =   "frmCases.frx":0039
            TabIndex        =   13
            Text            =   "PENDING"
            Top             =   2700
            Width           =   1455
         End
         Begin VB.TextBox txtClientID 
            DataField       =   "ClientID"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            TabIndex        =   8
            Top             =   1020
            Width           =   915
         End
         Begin VB.TextBox txtLawyerID 
            DataField       =   "LawyerID"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            TabIndex        =   10
            Top             =   1440
            Width           =   915
         End
         Begin VB.TextBox txtDescription 
            DataField       =   "Description"
            DataSource      =   "Data1"
            Height          =   585
            Left            =   1380
            MaxLength       =   49
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   1920
            Width           =   2235
         End
         Begin VB.TextBox txtID 
            DataField       =   "ID"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   915
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Bindings        =   "frmCases.frx":004E
            DataField       =   "CloseDate"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1380
            TabIndex        =   15
            Top             =   3660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20119553
            CurrentDate     =   39063
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "(Up to 50 characters)"
            ForeColor       =   &H8000000E&
            Height          =   435
            Left            =   360
            TabIndex        =   33
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Start Date:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   180
            TabIndex        =   28
            Top             =   3240
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Status:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Client ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Lawyer ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   1500
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Description:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   24
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Close Date:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   180
            TabIndex        =   22
            Top             =   3720
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraCLSearch 
      BackColor       =   &H8000000D&
      Caption         =   "Case Search"
      ForeColor       =   &H8000000E&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton btnPrintList2 
         Caption         =   "PRINT LIST"
         Height          =   495
         Left            =   8820
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmCases.frx":0070
         Height          =   1935
         Left            =   240
         OleObjectBlob   =   "frmCases.frx":0084
         TabIndex        =   19
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
         Left            =   4740
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblCases"
         Top             =   3240
         Width           =   1155
      End
      Begin VB.ComboBox cmbSearch 
         Height          =   315
         ItemData        =   "frmCases.frx":0A57
         Left            =   3540
         List            =   "frmCases.frx":0A6D
         TabIndex        =   1
         Text            =   "ID"
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
         TabIndex        =   18
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblSearch2 
         BackColor       =   &H8000000D&
         Caption         =   "Search by:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3900
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddnew_Click()

'-Adds a new Case record-

Data1.Recordset.AddNew
DTPicker1.Value = Date
DTPicker2.Value = Date + 365
btnClientS.Visible = True
btnLawyerS.Visible = True
btnShowClient.Visible = False
btnShowLawyer.Visible = False
btnClientS.SetFocus

'--------------------------

End Sub

Private Sub btnClientS_Click()

'-Calls Client Search form-

frmClientSearch.Show 1
If mdlMain.ClientID <> "" Then txtClientID.Text = mdlMain.ClientID
btnClientS.Visible = False

'--------------------------

End Sub

Private Sub btnDelete_Click()

'-Deletes selected Case record-

On Error GoTo mistake
X = MsgBox("The selected record will be permanently deleted. Are you sure?", vbOKCancel, "Warning!")
If X = vbOK Then
Data1.Recordset.Delete
Data1.Refresh
End If
mistake:

'--------------------------------

End Sub

Private Sub btnDocuments_Click()

'-Sends the CaseID value to the related variable in MdlMain-

Data1.Refresh
mdlMain.CaseID = txtID.Text
frmDocuments.Show 1

'-----------------------------------------------------------

End Sub

Private Sub btnLawyerS_Click()

'-Calls Lawyer Search form-

frmLawyerSearch.Show 1
If mdlMain.LawyerID <> "" Then txtLawyerID.Text = mdlMain.LawyerID
btnLawyerS.Visible = False

'--------------------------

End Sub

Private Sub btnPrint_Click()

'-Sends the selected Case details to the default printer----

On Error GoTo mistake
If txtID.Text <> "" Then
    Printer.Orientation = 1
    Printer.FontName = "Arial"
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print newline
        Printer.Print newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Font.Size = 16
        Printer.Print Tab(26); "GENESIS CASE DETAILS"; newline
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 12
        Printer.Print Tab(33); "ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Client ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtClientID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Lawyer ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtLawyerID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Status: ";
        Printer.FontBold = False
        Printer.Print Tab(53); cmbStatus.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Start Date: ";
        Printer.FontBold = False
        Printer.Print Tab(53); DTPicker1.Value;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "End Date: ";
        Printer.FontBold = False
        Printer.Print Tab(53); DTPicker2.Value;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Description: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtDescription.Text;
        Printer.FontBold = False
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print Tab(48); "Date "; Date;
        Printer.EndDoc
    End If
mistake:

'------------------------------------------------------------------

End Sub

Private Sub btnPrintList2_Click()

'-Sends the Cases list to the default printer----

On Error GoTo mistake
Dim counter As Byte
Dim pagecount As Byte
    counter = 0
    pagecount = 0
    Data1.Recordset.MoveFirst
    Data1.Refresh
    Printer.Orientation = 2
    Printer.FontName = "Arial"
    While Not Data1.Recordset.EOF
        counter = 0
        pagecount = pagecount + 1
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 16
        Printer.Print Tab(42); "GENESIS CASES REPORT"; newline
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 10
        Printer.Print Tab(15); "ID";
        Printer.Print Tab(28); "Client ID";
        Printer.Print Tab(56); "Lawyer ID";
        Printer.Print Tab(81); "Status";
        Printer.Print Tab(116); "Start Date";
        Printer.Print Tab(141); "End Date";
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        While counter < 17 And Not Data1.Recordset.EOF
            Printer.FontBold = False
            Printer.Print Tab(15); Data1.Recordset.ID;
            Printer.Print Tab(30); Data1.Recordset.ClientID;
            Printer.Print Tab(60); Data1.Recordset.LawyerID;
            Printer.Print Tab(88); Data1.Recordset.Status;
            Printer.Print Tab(126); Data1.Recordset.startdate;
            Printer.Print Tab(152); Data1.Recordset.closedate;
            Printer.Print newline
            Printer.Print newline
            Data1.Recordset.MoveNext
            counter = counter + 1
        Wend
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(16); "Date "; Date;
        Printer.Print Tab(152); "Page "; pagecount;
        Printer.FontBold = False
        Printer.NewPage
    Wend
    Printer.EndDoc
mistake:

'------------------------------------------------------------------

End Sub

Private Sub btnShowClient_Click()

'-Opens Client Details---------

mdlMain.ClientIDDetails = txtClientID.Text
frmClientDetails.Show 1

'-------------------------------

End Sub

Private Sub btnShowLawyer_Click()

'-Opens Lawyer Details---------

mdlMain.LawyerIDDetails = txtLawyerID.Text
frmLawyerDetails.Show 1

'-------------------------------

End Sub

Private Sub btnUpdate_Click()

'-Updates Case record-

On Error GoTo mistake
Data1.Refresh
If btnClientS.Visible = True Then btnClientS.Visible = False
If btnLawyerS.Visible = True Then btnLawyerS.Visible = False
If btnShowClient.Visible = False Then btnShowClient.Visible = True
If btnShowLawyer.Visible = False Then btnShowLawyer.Visible = True
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------------

End Sub


Private Sub Form_Load()

'-Loads Data to the form controls-----

On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "/dbGenesis.mdb"
Data1.Refresh
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'-------------------------------------

End Sub

Private Sub Form_Unload(Cancel As Integer)

'-Updates Case records-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'---------------------

End Sub

Private Sub txtName_Change()

'-Sends SQL commands to database. It works as a search engine for case records----------

On Error GoTo mistake:
If txtName.Text <> "" Then
    Select Case cmbSearch.Text
        Case "ID"
            Data1.RecordSource = "select * from tblCases where ID = " & txtName.Text & ";"
            Data1.Refresh
        Case Else
            Data1.RecordSource = "select * from tblCases where " & cmbSearch.Text & " like '" & txtName.Text & "*" & "';"
            Data1.Refresh
    End Select
Else
    Data1.RecordSource = "select * from tblCases;"
    Data1.Refresh
End If
Exit Sub

mistake:
MsgBox "Invalid input. Please insert a correct value.", vbExclamation, "Caution!"

'----------------------------------------------------------------------------------------

End Sub
