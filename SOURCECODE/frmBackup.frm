VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACKUP MODULE"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4995
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Backup Menu"
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4755
      Begin VB.CommandButton btnCopy 
         Caption         =   "BACKUP"
         Height          =   495
         Left            =   3060
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton btnReplace 
         Caption         =   "REPLACE"
         Height          =   495
         Left            =   3060
         TabIndex        =   1
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblSearch1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "Press to make a copy of current data:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Caption         =   "Press to replace current database:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1380
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   ".mdb"
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCopy_Click()

'-Makes a copy of the current database----

CommonDialog1.ShowSave
FileCopy mdlMain.ApplicPath & "/dbGENESIS.mdb", CommonDialog1.FileName & ".mdb"

'-----------------------------------------

End Sub


Private Sub btnReplace_Click()

'-Replaces current database with a older version----

MsgBox "Please choose the COPY of the data (NOT the ORIGINAL DATABASE)which will replace the Main Database file (dbGenesis.mdb).", vbOK, "WARNING!"
CommonDialog1.ShowOpen
X = MsgBox("Warning! All current data will be replaced by an older version. Are you sure?", vbOKCancel, "Replace Data?")
If X = vbOK Then
FileCopy CommonDialog1.FileName, mdlMain.ApplicPath & "/dbGENESIS.mdb"
End If

'----------------------------------------------------

End Sub


