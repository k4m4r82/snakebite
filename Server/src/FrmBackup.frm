VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Database"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2835
   Icon            =   "FrmBackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   2835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton CmdBackup 
      Caption         =   "Bakup"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Backup File Database"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBackup_Click()
    If FExists(App.path & "\" & BackupFileName) Then
        Kill App.path & "\" & BackupFileName
        FileCopy App.path & "\" & DataFileName, App.path & "\" & BackupFileName
    Else
        FileCopy App.path & "\" & DataFileName, App.path & "\" & BackupFileName
    End If
    Unload Me
End Sub

Private Sub CmdBackup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBackup, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtrlHover CmdBatal, X, Y, vbButtonFace, vbButtonShadow
End Sub

Private Sub Form_Load()
    CmdBatal.Caption = LangBatal
End Sub
