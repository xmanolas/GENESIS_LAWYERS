VERSION 5.00
Begin VB.Form frmLog 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WELCOME TO GENESIS - LOG IN"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Log In"
      ForeColor       =   &H8000000E&
      Height          =   1995
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   3555
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   900
         Width           =   1455
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "Username:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "Password:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtRight 
      DataField       =   "uRight"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      DataField       =   "passW"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtUser 
      DataField       =   "userN"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\GENESIS\dbUsers.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblUsers"
      Top             =   1740
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnOK_Click()

'-Searches if the inserted data belongs to a valid user and applies rights policy---

Data1.RecordSource = "select * from tblUsers where userN like '" & txtUsername.Text & "' and passW like '" & txtPassword.Text & "';"
Data1.Refresh
If txtUser.Text <> "" Then
mdlMain.UserName = txtUser.Text
mdlMain.Password = txtPass.Text
mdlMain.UserRight = txtRight.Text
Unload Me
frmMainMenu.Show 1
Else
MsgBox "Invalid Username or Password. Please try again.", vbCritical, "Invalid Data!"
End If

'-----------------------------------------------------------------------------------

End Sub


Private Sub Form_Load()

'-Loads Data to the form controls-----

mdlMain.paymentSwitch = False
On Error GoTo mistake
Data1.DatabaseName = mdlMain.ApplicPath & "\dbUsers.mdb"
Data1.Refresh
Exit Sub
mistake:
MsgBox "Application data directory is missing or corrupted! Please run again the GENESIS Application and open the Initialization Procedure (Press E in the Initialization screen).", vbInformation
End

'-------------------------------------

End Sub

