VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPayments 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT MANAGEMENT MODULE"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "frmPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Payments"
      ForeColor       =   &H8000000E&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9195
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000D&
         Caption         =   "Payment Details"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   180
         TabIndex        =   16
         Top             =   3480
         Width           =   6015
         Begin VB.CommandButton btnShowClient 
            Caption         =   "DETAILS"
            Height          =   315
            Left            =   2760
            TabIndex        =   9
            Top             =   780
            Width           =   1035
         End
         Begin VB.TextBox txtClientID 
            DataField       =   "ClientID"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   49
            TabIndex        =   8
            Top             =   780
            Width           =   915
         End
         Begin VB.TextBox txtPayID 
            DataField       =   "ID"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   49
            TabIndex        =   7
            Top             =   300
            Width           =   915
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "PaymentDate"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            TabIndex        =   11
            Top             =   1740
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   151781377
            CurrentDate     =   38718
         End
         Begin VB.TextBox txtDescription 
            DataField       =   "Description"
            DataSource      =   "Data1"
            Height          =   555
            Left            =   1740
            MaxLength       =   49
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2220
            Width           =   2955
         End
         Begin VB.TextBox txtAmount 
            DataField       =   "Amount"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1740
            MaxLength       =   49
            TabIndex        =   10
            Top             =   1260
            Width           =   1395
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "(Up to 50 characters)"
            ForeColor       =   &H8000000E&
            Height          =   435
            Left            =   660
            TabIndex        =   25
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Client ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   420
            TabIndex        =   24
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Payment ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   420
            TabIndex        =   20
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Date:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   720
            TabIndex        =   19
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Description:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   720
            TabIndex        =   18
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Amount:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Controls"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   6360
         TabIndex        =   14
         Top             =   3480
         Width           =   2655
         Begin VB.CommandButton btnPrint 
            Caption         =   "PRINT"
            Height          =   435
            Left            =   600
            TabIndex        =   6
            Top             =   2220
            Width           =   1455
         End
         Begin VB.CommandButton btnUpdate 
            Caption         =   "UPDATE"
            Height          =   435
            Left            =   600
            TabIndex        =   4
            Top             =   1020
            Width           =   1455
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "DELETE"
            Height          =   435
            Left            =   600
            TabIndex        =   5
            Top             =   1620
            Width           =   1455
         End
         Begin VB.CommandButton btnAdd 
            Caption         =   "NEW"
            Height          =   435
            Left            =   600
            TabIndex        =   3
            Top             =   420
            Width           =   1455
         End
      End
      Begin VB.Frame fraCLSearch 
         BackColor       =   &H8000000D&
         Caption         =   "Payment History"
         ForeColor       =   &H8000000E&
         Height          =   3075
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   8835
         Begin VB.CommandButton btnShowAll 
            Caption         =   "SHOW ALL PAYMENTS"
            Height          =   555
            Left            =   4800
            TabIndex        =   2
            Top             =   300
            Width           =   1635
         End
         Begin VB.CommandButton btnClientSearch 
            Caption         =   "SEARCH BY CLIENT"
            Height          =   555
            Left            =   2640
            TabIndex        =   1
            Top             =   300
            Width           =   1635
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
            Left            =   3960
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblPayments"
            Top             =   2580
            Width           =   1155
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmPayments.frx":000C
            Height          =   1335
            Left            =   240
            OleObjectBlob   =   "frmPayments.frx":0020
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1020
            Width           =   8355
         End
         Begin VB.Label lblClientFirstname 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   2400
            TabIndex        =   22
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label lblClientSurname 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4500
            TabIndex        =   21
            Top             =   300
            Width           =   2055
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
      TabIndex        =   23
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "frmPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClientSearch_Click()

'-Calls Client Search form-

frmClientSearch.Show 1
On Error GoTo mistake
If mdlMain.ClientID <> "" Then
lblClientID.Caption = mdlMain.ClientID
lblClientFirstname.Caption = mdlMain.ClientFirstname
lblClientSurname.Caption = mdlMain.ClientSurname
Data1.RecordSource = "select * from tblPayments where ClientID like '" & lblClientID.Caption & "';"
Data1.Refresh
End If
Exit Sub
mistake:
MsgBox "An error occured during data loading. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------

End Sub

Private Sub btnPrint_Click()

'-Sends the selected Payment details to the default printer----

On Error GoTo mistake
    If txtPayID.Text <> "" Then
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
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Font.Size = 16
        Printer.Print Tab(29); "PAYMENT DETAILS"; newline
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 12
        Printer.Print Tab(33); "Payment ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtPayID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Client ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtClientID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Amount: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtAmount.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Date: ";
        Printer.FontBold = False
        Printer.Print Tab(53); DTPicker1.Value;
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

Private Sub btnDelete_Click()

'-Deletes selected Payment record-

On Error GoTo mistake
X = MsgBox("The selected record will be permanently deleted. Are you sure?", vbOKCancel, "Warning!")
If X = vbOK Then
Data1.Recordset.Delete
Data1.Refresh
End If
mistake:

'--------------------------------

End Sub

Private Sub btnShowAll_Click()

'-Shows all payments from all Clients------------

Data1.RecordSource = "select * from tblPayments;"
Data1.Refresh

'------------------------------------------------

End Sub

Private Sub btnShowClient_Click()

'-Opens Client Details---------

mdlMain.ClientIDDetails = txtClientID.Text
frmClientDetails.Show 1

'-------------------------------

End Sub

Private Sub btnUpdate_Click()

'-Updates selected Client record-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data loading. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------------

End Sub


Private Sub btnAdd_Click()

'-Adds a new Payment record for the selected Client-

On Error GoTo mistake
If lblClientID.Caption <> "" Then
Data1.Recordset.AddNew
DTPicker1.Value = Date
txtClientID.Text = lblClientID.Caption
txtAmount.SetFocus
Else
MsgBox "Please select a Client before registering the related Payment.", vbExclamation, "No Client Selected!"
End If
Exit Sub
mistake:
MsgBox "An error occured during data loading. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'----------------------------------------------------

End Sub




Private Sub Form_Load()

'-Tries to read and load Client data values from related variable in mdlMain-

If mdlMain.paymentSwitch = True Then btnAdd.Enabled = True Else btnAdd.Enabled = False
On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "/dbGenesis.mdb"
If mdlMain.ClientID <> "" Then
lblClientID.Caption = mdlMain.ClientID
lblClientFirstname.Caption = mdlMain.ClientFirstname
lblClientSurname.Caption = mdlMain.ClientSurname
Data1.RecordSource = "select * from tblPayments where ClientID like '" & lblClientID.Caption & "';"
Data1.Refresh
Else
Data1.RecordSource = "select * from tblPayments;"
Data1.Refresh
End If
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'----------------------------------------------------------------------------

End Sub



Private Sub Form_Unload(Cancel As Integer)

mdlMain.paymentSwitch = False

End Sub
