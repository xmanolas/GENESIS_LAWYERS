VERSION 5.00
Begin VB.Form frmLawyerDetails 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAWYER DETAILS"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\GENESIS\dbGenesis.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblLawyers"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000D&
      DataField       =   "Email"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   17
      Top             =   3600
      Width           =   2475
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "E-mail:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000D&
      DataField       =   "Surname"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   15
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      DataField       =   "Firstname"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   14
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000D&
      DataField       =   "Address"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   13
      Top             =   1500
      Width           =   2475
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      DataField       =   "PostalCode"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      DataField       =   "Country"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   11
      Top             =   2340
      Width           =   1515
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      DataField       =   "Tel1"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   10
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000D&
      DataField       =   "Tel2"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   9
      Top             =   3180
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      DataField       =   "ID"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2580
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Surname:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Firstname:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1140
      TabIndex        =   6
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Address"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Postal Code:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Country:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Tel 1:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "Tel 2:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "ID:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmLawyerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'-Loads current Lawyer's details---------

On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "/dbGenesis.mdb"
Data1.Refresh
If mdlMain.LawyerIDDetails <> "" Then
Data1.RecordSource = "select * from tblLawyers where ID = " & mdlMain.LawyerIDDetails & ";"
Data1.Refresh
End If
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'----------------------------------------

End Sub

