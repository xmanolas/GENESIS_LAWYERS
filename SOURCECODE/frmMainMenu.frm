VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENESIS LEGAL SERVICES"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMainMenu 
      BackColor       =   &H8000000D&
      Caption         =   "MAIN MENU"
      ForeColor       =   &H8000000E&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton btnLawyers 
         Caption         =   "LAWYERS"
         Height          =   495
         Left            =   3060
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton btnAdmin 
         Caption         =   "ADMIN"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "HELP"
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton btnBackup 
         Caption         =   "BACKUP"
         Height          =   495
         Left            =   2100
         TabIndex        =   5
         Top             =   3540
         Width           =   1455
      End
      Begin VB.CommandButton btnPayments 
         Caption         =   "PAYMENTS"
         Height          =   495
         Left            =   3060
         TabIndex        =   4
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CommandButton btnCases 
         Caption         =   "CASES"
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CommandButton btnClients 
         Caption         =   "CLIENTS"
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "SECURITY LEVEL :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H8000000D&
         DataField       =   "ID"
         DataSource      =   "Data1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2940
         TabIndex        =   10
         Top             =   780
         Width           =   2475
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "USERNAME :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H8000000D&
         DataField       =   "ID"
         DataSource      =   "Data1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2940
         TabIndex        =   8
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.OLE OLE1 
      Class           =   "Word.Document.8"
      Height          =   1035
      Left            =   2100
      OleObjectBlob   =   "frmMainMenu.frx":0000
      SourceDoc       =   "C:\Program Files\GENESIS\Help.rtf"
      TabIndex        =   12
      Top             =   4260
      Width           =   1035
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdmin_Click()

'-Opens User Administration Module-

frmUsers.Show 1

'---------------------------------

End Sub

Private Sub btnBackup_Click()

'-Opens Backup Module-

frmBackup.Show 1

'---------------------

End Sub

Private Sub btnCases_Click()

'-Opens Cases Management Module-

frmCases.Show 1

'-------------------------------

End Sub

Private Sub btnClientS_Click()

'-Opens Client Registration Module-

frmClientReg.Show 1

'----------------------------------

End Sub

Private Sub btnHelp_Click()

'-Opens the Help File---

On Error GoTo mistake
OLE1.CreateLink (mdlMain.ApplicPath & "\Help.rtf")
OLE1.DoVerb (1)
Exit Sub
mistake:
MsgBox "Help File missing or Corrupted! Please refer either to ReadMe.doc or the Help.rtf file on program's installation directory or contact the SMALL VILLE HELPDESK.", vbInformation, "File Missing!"

'-----------------------

End Sub

Private Sub btnLawyerS_Click()

'-Opens Lawyer Registration Module-

frmLawyers.Show 1

'----------------------------------

End Sub

Private Sub btnPayments_Click()

'-Opens Payments Management Module-

mdlMain.paymentSwitch = False
frmPayments.Show 1

'------------------------------

End Sub



Private Sub Form_Activate()

'-Checks user rights and displays the relevant buttons-----

Select Case mdlMain.UserRight
    Case "User"
        btnBackup.Enabled = False
        btnPayments.Enabled = False
        btnAdmin.Enabled = False
    Case "Financial Manager"
        btnCases.Enabled = False
        btnBackup.Enabled = False
        btnAdmin.Enabled = False
        btnLawyerS.Enabled = False
    Case "Power User"
        btnAdmin.Enabled = False
End Select
lblUsername.Caption = mdlMain.UserName
lblPassword.Caption = mdlMain.UserRight
        
'-----------------------------------------------------------

End Sub

