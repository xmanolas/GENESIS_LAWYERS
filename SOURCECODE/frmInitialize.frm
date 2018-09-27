VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInitialize 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " INITIALIZATION MODULE"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   2580
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   2235
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   3915
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   435
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   1035
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   600
         Width           =   1875
      End
      Begin VB.DirListBox Dir1 
         Height          =   990
         Left            =   300
         TabIndex        =   7
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Please select the GENESIS installation directory:"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   3675
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3060
      Top             =   3420
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\GENESIS\dbInitialize.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblPath"
      Top             =   3900
      Width           =   1875
   End
   Begin VB.TextBox txtPath 
      DataField       =   "mainPath"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   3915
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   6
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Press letter "" i "" on the keyboard to enter the Initialization Menu"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3315
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "6"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1860
         TabIndex        =   4
         Top             =   1140
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()

'-Updates Data Directory location----

txtPath.Text = Dir1.Path
Data1.Refresh
Timer1.Interval = 1000
Frame2.Visible = False

'------------------------------------

End Sub

Private Sub Drive1_Change()

'-Changes the disk drive location--------

On Error GoTo mistake
Dir1.Path = Drive1.Drive
Exit Sub

mistake:
MsgBox "Pease select a valid Disk Drive letter!", vbInformation, "Warning!"

'----------------------------------------

End Sub

Private Sub Form_Load()

'-Loads the default Application Path Data---

Data1.DatabaseName = App.Path & "/dbInitialize.mdb"
Data1.Refresh

'-------------------------------------------

End Sub


Private Sub IniProcedure()

'-Shows the Initialization Window-----------
    
MsgBox "You are about to run GENESIS COMPUTERIZED SYSTEM initialization procedure and to set the program's data directory location. Please, in the following Directory Navigation Controls indicate the location of GENESIS installation folder and press OK button", vbInformation
Frame2.Visible = True
Timer1.Interval = 0

'-------------------------------------------

End Sub



Private Sub Form_Unload(Cancel As Integer)

'-Loads selected Application Path to the Program's Registry-----

If txtPath.Text = "NEW" Then txtPath.Text = App.Path
mdlMain.ApplicPath = txtPath.Text

'---------------------------------------------------------------

End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

'-Checks if the user pressed "I" button and (if so) calls the Initialization procedure-----

If KeyCode = vbKeyI Then Call IniProcedure

'------------------------------------------------------------------------------------------

End Sub

Private Sub Timer1_Timer()

'-Counts the time that passes before loading the Log In form (5 seconds)---

If ProgressBar1.Value < 6 Then
ProgressBar1.Value = ProgressBar1.Value + 1
Label2.Caption = 6 - ProgressBar1.Value
Else
Unload Me
frmLog.Show 1
End If

'--------------------------------------------------------------------------

End Sub


