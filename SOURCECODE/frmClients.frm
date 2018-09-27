VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClientReg 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENT REGISTRATION MODULE"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "frmClients.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraCLUpdate 
      BackColor       =   &H8000000D&
      Caption         =   "Client Information"
      ForeColor       =   &H8000000E&
      Height          =   5235
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   10695
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Controls"
         ForeColor       =   &H8000000E&
         Height          =   4875
         Left            =   5340
         TabIndex        =   31
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton btnPrint 
            Caption         =   "PRINT"
            Height          =   495
            Left            =   1920
            TabIndex        =   6
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "DELETE"
            Height          =   495
            Left            =   1920
            TabIndex        =   5
            Top             =   2700
            Width           =   1455
         End
         Begin VB.CommandButton btnUpdate 
            Caption         =   "UPDATE"
            Height          =   495
            Left            =   1920
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton btnAddnew 
            Caption         =   "CREATE"
            Height          =   495
            Left            =   1920
            TabIndex        =   3
            Top             =   900
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000D&
         Caption         =   "Details"
         ForeColor       =   &H8000000E&
         Height          =   4875
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton btnPayments 
            Caption         =   "PAYMENTS"
            Height          =   495
            Left            =   1620
            TabIndex        =   16
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox txtCountry 
            DataField       =   "Country"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   12
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtPostal 
            DataField       =   "PostalCode"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   11
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox txtSurname 
            DataField       =   "Surname"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   8
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtFirstname 
            DataField       =   "Firstname"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   9
            Top             =   1140
            Width           =   2055
         End
         Begin VB.TextBox txtAddress 
            DataField       =   "Address"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   10
            Top             =   1560
            Width           =   3195
         End
         Begin VB.TextBox txtID 
            DataField       =   "ID"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   49
            TabIndex        =   7
            Top             =   300
            Width           =   915
         End
         Begin VB.TextBox txtTel2 
            DataField       =   "Tel2"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   14
            Top             =   3240
            Width           =   2055
         End
         Begin VB.TextBox txtEmail 
            DataField       =   "Email"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   15
            Top             =   3660
            Width           =   2055
         End
         Begin VB.TextBox txtTel1 
            DataField       =   "Tel1"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1380
            MaxLength       =   49
            TabIndex        =   13
            Top             =   2820
            Width           =   2055
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Country:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Postal Code:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   29
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Surname:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   28
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Firstname:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   27
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Address:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "ID:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "E-mail:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   24
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Phone 2:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   3300
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Caption         =   "Phone 1:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   2880
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraCLSearch 
      BackColor       =   &H8000000D&
      Caption         =   "Client Search"
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
         Bindings        =   "frmClients.frx":000C
         Height          =   1935
         Left            =   240
         OleObjectBlob   =   "frmClients.frx":0020
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
         RecordSource    =   "tblClients"
         Top             =   3240
         Width           =   1155
      End
      Begin VB.ComboBox cmbSearch 
         Height          =   315
         ItemData        =   "frmClients.frx":09F3
         Left            =   3540
         List            =   "frmClients.frx":0A12
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
Attribute VB_Name = "frmClientReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddnew_Click()

'-Adds a new Client record-

Data1.Recordset.AddNew
txtSurname.SetFocus

'--------------------------

End Sub

Private Sub btnDelete_Click()

'-Deletes selected Client record-

On Error GoTo mistake
X = MsgBox("The selected record will be permanently deleted. Are you sure?", vbOKCancel, "Warning!")
If X = vbOK Then
Data1.Recordset.Delete
Data1.Refresh
End If
Exit Sub
mistake:

'--------------------------------

End Sub

Private Sub btnPayments_Click()

'-Opens current Client's Payments form  --

mdlMain.ClientID = txtID.Text
mdlMain.ClientFirstname = txtFirstname.Text
mdlMain.ClientSurname = txtSurname.Text
mdlMain.paymentSwitch = True
frmPayments.Show 1

'-----------------------------------------

End Sub

Private Sub btnPrint_Click()

'-Sends the selected Client details to the default printer----

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
        Printer.Print Tab(25); "GENESIS CLIENT DETAILS"; newline
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 12
        Printer.Print Tab(33); "ID: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtID.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Surname: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtSurname.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Firstname: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtFirstname.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Address: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtAddress.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Postal Code: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtPostal.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Country: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtCountry.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Tel1: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtTel1.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "Tel2: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtTel2.Text;
        Printer.FontBold = True
        Printer.Print newline
        Printer.Print newline
        Printer.Print Tab(33); "E-mail: ";
        Printer.FontBold = False
        Printer.Print Tab(53); txtEmail.Text;
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print Tab(48); "Date "; Date;
        Printer.FontBold = False
        Printer.EndDoc
    End If
mistake:
'------------------------------------------------------------------

End Sub

Private Sub btnPrintList2_Click()

'-Sends the Clients list to the default printer----

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
        Printer.Print Tab(42); "GENESIS CLIENTS REPORT"; newline
        Printer.Print newline
        Printer.Print newline
        Printer.Font.Size = 10
        Printer.Print Tab(15); "ID";
        Printer.Print Tab(28); "Surname";
        Printer.Print Tab(56); "Firstname";
        Printer.Print Tab(81); "Address";
        Printer.Print Tab(116); "Tel";
        Printer.Print Tab(141); "E-mail";
        Printer.Print ; newline
        Printer.Print ; newline
        Printer.Print ; newline
        While counter < 17 And Not Data1.Recordset.EOF
            Printer.FontBold = False
            Printer.Print Tab(15); Data1.Recordset.ID;
            Printer.Print Tab(30); Data1.Recordset.Surname;
            Printer.Print Tab(60); Data1.Recordset.Firstname;
            Printer.Print Tab(88); Data1.Recordset.Address;
            Printer.Print Tab(126); Data1.Recordset.Tel1;
            Printer.Print Tab(152); Data1.Recordset.email;
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

Private Sub btnUpdate_Click()

'-Updates selected Client record-

On Error GoTo mistake
Data1.Refresh
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

'-Updates selected Client record-

On Error GoTo mistake
Data1.Refresh
Exit Sub
mistake:
MsgBox "An error occured during data update. Information may not be stored correctly. Please close the form and try to re-open it.", vbExclamation, "Caution!"

'--------------------------------

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
